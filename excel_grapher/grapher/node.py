from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from enum import StrEnum
from types import MappingProxyType
from typing import Any, Protocol, TypeAlias

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    CellKey,
    NodeShape,
    ParsedRowKey,
    RangeKey,
    UnionKey,
)
from excel_grapher.core.address_keys import NodeKey as AddressKey
from excel_grapher.core.address_keys import expand_node_cells as _expand_node_cells
from excel_grapher.core.address_keys import format_cell_key as _format_cell_key
from excel_grapher.core.address_keys import format_row_key as _format_row_key
from excel_grapher.core.address_keys import members_to_node_key as _members_to_node_key
from excel_grapher.core.address_keys import normalize_row_key as _normalize_row_key
from excel_grapher.core.address_keys import parse_address as _parse_address
from excel_grapher.core.address_keys import parse_node_key as _parse_node_key
from excel_grapher.core.address_keys import parse_row_key as _parse_row_key
from excel_grapher.core.formula_ast import AstNode
from excel_grapher.grapher.formula_groups import AddressLeaf, validate_group_template

# Graph node identity; same canonical form as NormalizedAddress (cell / range / union).
NodeKey: TypeAlias = str


class NodeKind(StrEnum):
    """Compatibility kind for graph nodes.

    Prefer `Node.shape` / `parse_node_key(node.address)`. `row` remains for one-row
    `RangeKey` shims; other multi-cell nodes use `union`.
    """

    cell = "cell"
    row = "row"
    union = "union"


def _kind_for_address(address: AddressKey) -> NodeKind:
    if isinstance(address, CellKey):
        return NodeKind.cell
    if isinstance(address, RangeKey) and address.shape is NodeShape.row:
        return NodeKind.row
    return NodeKind.union


def _coerce_address(value: str | AddressKey) -> AddressKey:
    """Canonicalize `value`, preserving explicit 1x1 `RangeKey` row shims."""
    if isinstance(value, RangeKey):
        # Keep `Sheet1!D63:D63` distinct from `Sheet1!D63` for make_row_node shims.
        return value
    if isinstance(value, (CellKey, UnionKey)):
        return _parse_node_key(str(value))
    return _parse_node_key(value)


@dataclass
class Node:
    """Workbook cell or multi-cell span in a dependency graph.

    Canonical identity is `address` (`CellKey` / `RangeKey` / `UnionKey`). Legacy
    constructors that pass `sheet`/`column`/`row` (or row extent fields) still
    work and sync `address` in `__post_init__`.

    `is_leaf` is true when the node has no outgoing dependency edges (value-only
    cells and literal-only formulas such as `=1+1`).
    """

    sheet: str | None
    column: str | None
    row: int | None
    formula: str | None
    normalized_formula: str | None
    value: Any
    is_leaf: bool
    is_target: bool = False
    metadata: dict[str, Any] = field(default_factory=dict)
    kind: NodeKind = NodeKind.cell
    min_col: str | None = None
    min_row: int | None = None
    max_col: str | None = None
    max_row: int | None = None
    address: AddressKey | None = None
    shape_fingerprint: str | None = None
    skeleton: AstNode | None = None
    member_bindings: dict[str, tuple[AddressLeaf, ...]] | None = None

    def __post_init__(self) -> None:
        if self.address is not None:
            self.address = _coerce_address(self.address)
            self._sync_fields_from_address()
            self._finalize_template_fields()
            return

        if self.kind is NodeKind.row or (
            self.min_col is not None
            and self.max_col is not None
            and self.min_row is not None
            and self.max_row is not None
            and (self.column is None or self.kind is NodeKind.row)
        ):
            self._finalize_row_legacy()
            assert self.min_col is not None and self.max_col is not None
            assert self.min_row is not None and self.row is not None
            assert self.sheet is not None
            self.address = RangeKey(
                _format_row_key(self.sheet, self.min_col, self.row, self.max_col)
            )
            self.kind = NodeKind.row
            self.value = None
            self._finalize_template_fields()
            return

        self._finalize_cell_legacy()
        assert self.sheet is not None and self.column is not None and self.row is not None
        self.address = CellKey(_format_cell_key(self.sheet, self.column, self.row))
        self.kind = NodeKind.cell
        self._finalize_template_fields()

    def _finalize_template_fields(self) -> None:
        """Validate or clear formula-group template fields after address sync."""
        has_any = (
            self.shape_fingerprint is not None
            or self.skeleton is not None
            or self.member_bindings is not None
        )
        if not has_any:
            self.shape_fingerprint = None
            self.skeleton = None
            self.member_bindings = None
            return

        if self.kind is NodeKind.cell or (
            self.address is not None and self.address.shape is NodeShape.cell
        ):
            raise ValueError("Formula-group template fields are only allowed on multi-cell nodes")

        if self.skeleton is None or self.member_bindings is None or self.shape_fingerprint is None:
            raise ValueError(
                "Formula-group template requires shape_fingerprint, skeleton, and member_bindings"
            )

        members = member_keys(self)
        self.member_bindings = validate_group_template(
            members=members,
            skeleton=self.skeleton,
            member_bindings=self.member_bindings,
            shape_fingerprint_value=self.shape_fingerprint,
        )

    def _sync_fields_from_address(self) -> None:
        addr = self.address
        assert addr is not None
        self.kind = _kind_for_address(addr)
        if isinstance(addr, CellKey):
            self.sheet = addr.sheet
            self.column = addr.column
            self.row = addr.row
            self.min_col = addr.column
            self.max_col = addr.column
            self.min_row = addr.row
            self.max_row = addr.row
            return

        if isinstance(addr, RangeKey):
            self.sheet = addr.sheet
            self.min_col = addr.min_col
            self.max_col = addr.max_col
            self.min_row = addr.min_row
            self.max_row = addr.max_row
            self.column = addr.column
            self.row = addr.row
            self.value = None
            return

        # UnionKey
        sheets = {m.sheet for m in addr.members}
        self.sheet = next(iter(sheets)) if len(sheets) == 1 else None
        self.column = None
        self.row = None
        self.min_col = None
        self.max_col = None
        self.min_row = None
        self.max_row = None
        self.value = None

    def _finalize_cell_legacy(self) -> None:
        if self.sheet is None or self.column is None or self.row is None:
            raise ValueError("Cell nodes require sheet, column, and row")
        col = str(self.column).upper()
        row = int(self.row)
        self.column = col
        self.row = row
        if self.min_col is None:
            self.min_col = col
        if self.max_col is None:
            self.max_col = col
        if self.min_row is None:
            self.min_row = row
        if self.max_row is None:
            self.max_row = row
        self.min_col = str(self.min_col).upper()
        self.max_col = str(self.max_col).upper()
        self.min_row = int(self.min_row)
        self.max_row = int(self.max_row)

    def _finalize_row_legacy(self) -> None:
        if self.sheet is None:
            raise ValueError("Row nodes require sheet")
        if self.min_col is None or self.max_col is None:
            raise ValueError("Row nodes require min_col and max_col")
        if self.min_row is None or self.max_row is None:
            raise ValueError("Row nodes require min_row and max_row")
        min_row = int(self.min_row)
        max_row = int(self.max_row)
        if min_row != max_row:
            raise ValueError(
                f"Row nodes must be a one-row extent (min_row == max_row); "
                f"got min_row={min_row}, max_row={max_row}"
            )
        if self.row is not None and int(self.row) != min_row:
            raise ValueError(
                f"Row node row must match min_row/max_row; got row={self.row}, min_row={min_row}"
            )
        left_idx = fastpyxl.utils.cell.column_index_from_string(str(self.min_col))
        right_idx = fastpyxl.utils.cell.column_index_from_string(str(self.max_col))
        if left_idx <= right_idx:
            min_col = str(self.min_col).upper()
            max_col = str(self.max_col).upper()
        else:
            min_col = str(self.max_col).upper()
            max_col = str(self.min_col).upper()
        self.column = None
        self.row = min_row
        self.min_col = min_col
        self.max_col = max_col
        self.min_row = min_row
        self.max_row = max_row
        self.kind = NodeKind.row

    @property
    def key(self) -> NodeKey:
        assert self.address is not None
        return str(self.address)

    @property
    def shape(self) -> NodeShape:
        assert self.address is not None
        return self.address.shape

    @property
    def column_index(self) -> int:
        if self.kind is not NodeKind.cell or self.column is None:
            raise ValueError("column_index is only defined for cell nodes")
        return int(fastpyxl.utils.cell.column_index_from_string(self.column))


@dataclass(frozen=True)
class NodeView:
    """Read-only snapshot of a graph node.

    Instances are returned by `DependencyGraph.get_node(...)` and reflect the
    node's state at the time of the lookup. To observe subsequent mutations,
    re-fetch the view. Durable mutation is done via
    `DependencyGraph.set_node_value(...)` and `set_node_metadata(...)`.
    """

    sheet: str | None
    column: str | None
    row: int | None
    formula: str | None
    normalized_formula: str | None
    value: Any
    is_leaf: bool
    is_target: bool
    metadata: Mapping[str, Any]
    kind: NodeKind = NodeKind.cell
    min_col: str | None = None
    min_row: int | None = None
    max_col: str | None = None
    max_row: int | None = None
    address: AddressKey | None = None
    shape_fingerprint: str | None = None
    skeleton: AstNode | None = None
    member_bindings: Mapping[str, tuple[AddressLeaf, ...]] | None = None

    @property
    def key(self) -> NodeKey:
        if self.address is not None:
            return str(self.address)
        if self.kind is NodeKind.row:
            assert self.sheet is not None
            assert self.min_col is not None and self.max_col is not None and self.row is not None
            return _format_row_key(self.sheet, self.min_col, self.row, self.max_col)
        assert self.sheet is not None and self.column is not None and self.row is not None
        return _format_cell_key(self.sheet, self.column, self.row)

    @property
    def shape(self) -> NodeShape:
        if self.address is not None:
            return self.address.shape
        if self.kind is NodeKind.row:
            return NodeShape.row
        if self.kind is NodeKind.union:
            return NodeShape.union
        return NodeShape.cell

    @property
    def column_index(self) -> int:
        if self.kind is not NodeKind.cell or self.column is None:
            raise ValueError("column_index is only defined for cell nodes")
        return int(fastpyxl.utils.cell.column_index_from_string(self.column))


def node_to_view(node: Node) -> NodeView:
    """Build an immutable `NodeView` snapshot from a stored `Node`."""
    bindings: Mapping[str, tuple[AddressLeaf, ...]] | None
    if node.member_bindings is None:
        bindings = None
    else:
        bindings = MappingProxyType(dict(node.member_bindings))
    return NodeView(
        sheet=node.sheet,
        column=node.column,
        row=node.row,
        formula=node.formula,
        normalized_formula=node.normalized_formula,
        value=node.value,
        is_leaf=node.is_leaf,
        is_target=node.is_target,
        metadata=MappingProxyType(dict(node.metadata)),
        kind=node.kind,
        min_col=node.min_col,
        min_row=node.min_row,
        max_col=node.max_col,
        max_row=node.max_row,
        address=node.address,
        shape_fingerprint=node.shape_fingerprint,
        skeleton=node.skeleton,
        member_bindings=bindings,
    )


def make_cell_node(
    sheet: str,
    column: str,
    row: int,
    *,
    formula: str | None = None,
    normalized_formula: str | None = None,
    value: Any = None,
    is_leaf: bool = True,
    is_target: bool = False,
    metadata: dict[str, Any] | None = None,
) -> Node:
    """Build a single-cell graph node."""
    return Node(
        sheet=sheet,
        column=column,
        row=row,
        formula=formula,
        normalized_formula=normalized_formula,
        value=value,
        is_leaf=is_leaf,
        is_target=is_target,
        metadata=dict(metadata or {}),
        kind=NodeKind.cell,
    )


def make_union_node(
    members: Sequence[str | AddressKey],
    *,
    formula: str | None = None,
    normalized_formula: str | None = None,
    value: Any = None,
    is_leaf: bool = True,
    is_target: bool = False,
    metadata: dict[str, Any] | None = None,
    shape_fingerprint: str | None = None,
    skeleton: AstNode | None = None,
    member_bindings: Mapping[str, Sequence[AddressLeaf]] | None = None,
) -> Node:
    """Build a multi-cell node from cell members (or a cell node for one member).

    `value` is ignored for multi-cell nodes (always stored as `None`). A single
    member collapses to a cell node via `members_to_node_key`.

    Optional formula-group template fields (`shape_fingerprint`, `skeleton`,
    `member_bindings`) are validated together and only allowed on multi-cell
    nodes.
    """
    if not members:
        raise ValueError("Cannot build node from empty member set")
    address = _members_to_node_key(members)
    bindings_dict: dict[str, tuple[AddressLeaf, ...]] | None
    if member_bindings is None:
        bindings_dict = None
    else:
        bindings_dict = {str(k): tuple(v) for k, v in member_bindings.items()}

    if isinstance(address, CellKey):
        if shape_fingerprint is not None or skeleton is not None or bindings_dict is not None:
            raise ValueError("Formula-group template fields are only allowed on multi-cell nodes")
        return Node(
            sheet=None,
            column=None,
            row=None,
            formula=formula,
            normalized_formula=normalized_formula,
            value=value,
            is_leaf=is_leaf,
            is_target=is_target,
            metadata=dict(metadata or {}),
            address=address,
        )
    return Node(
        sheet=None,
        column=None,
        row=None,
        formula=formula,
        normalized_formula=normalized_formula,
        value=None,
        is_leaf=is_leaf,
        is_target=is_target,
        metadata=dict(metadata or {}),
        address=address,
        shape_fingerprint=shape_fingerprint,
        skeleton=skeleton,
        member_bindings=bindings_dict,
    )


def make_row_node(
    sheet: str,
    row: int,
    min_col: str,
    max_col: str,
    *,
    formula: str | None = None,
    normalized_formula: str | None = None,
    value: Any = None,
    is_leaf: bool = True,
    is_target: bool = False,
    metadata: dict[str, Any] | None = None,
) -> Node:
    """Build a one-row graph node for `sheet!min_col{row}:max_col{row}`.

    Deprecated compatibility shim over a `RangeKey` address (including 1x1
    `D63:D63`, distinct from a cell node). Prefer `make_union_node` or
    `Node(..., address=...)` for new code.

    Raises:
        ValueError: If the resulting extent is not a single worksheet row.
    """
    key = _format_row_key(sheet, min_col, row, max_col)
    return Node(
        sheet=sheet,
        column=None,
        row=row,
        formula=formula,
        normalized_formula=normalized_formula,
        value=None,
        is_leaf=is_leaf,
        is_target=is_target,
        metadata=dict(metadata or {}),
        kind=NodeKind.row,
        min_col=min_col,
        min_row=row,
        max_col=max_col,
        max_row=row,
        address=RangeKey(key),
    )


def member_keys(node: Node | NodeView) -> list[NodeKey]:
    """Return sheet-qualified cell keys owned by `node` (expansion of `address`)."""
    if node.address is None:
        raise ValueError("Node is missing address")
    return [str(c) for c in _expand_node_cells(node.address)]


def row_member_keys(node: Node | NodeView) -> list[NodeKey]:
    """Return sheet-qualified cell keys for each column in a row node's span.

    Members share sheet and row; only the column differs. Raises `ValueError`
    when `node` is not a one-row node.
    """
    if node.kind is not NodeKind.row and node.shape is not NodeShape.row:
        raise ValueError("row_member_keys requires a row node")
    return member_keys(node)


def members_differ_only_by_column(keys: Sequence[NodeKey]) -> bool:
    """Return True when every key shares sheet+row and columns are unique."""
    if not keys:
        return True

    sheets: set[str] = set()
    rows: set[int] = set()
    cols: set[str] = set()
    for key in keys:
        sheet, coord = _parse_address(key)
        col, row = fastpyxl.utils.cell.coordinate_from_string(coord)
        sheets.add(sheet)
        rows.add(int(row))
        cols.add(str(col).upper())
    return len(sheets) == 1 and len(rows) == 1 and len(cols) == len(keys)


@dataclass(frozen=True, slots=True)
class CellLocation:
    """Where a single workbook cell is represented in a dependency graph.

    Attributes:
        cell_key: Canonical single-cell key that was queried.
        kind: Owning node kind (`cell`, `row`, or `union`).
        node_key: Graph key of the owning node.
        column: Column letters of the queried cell.
    """

    cell_key: NodeKey
    kind: NodeKind
    node_key: NodeKey
    column: str


@dataclass(frozen=True, slots=True)
class RangeLocation:
    """Where a one-row range is represented in a dependency graph.

    Attributes:
        range_key: Canonical one-row query key (after column ordering).
        kind: Owning node kind (`row` for exact or covering row nodes).
        node_key: Graph key of the owning row node.
        min_col: Query span left column.
        max_col: Query span right column.
        row: Query worksheet row.
    """

    range_key: NodeKey
    kind: NodeKind
    node_key: NodeKey
    min_col: str
    max_col: str
    row: int


def _require_single_cell_key(cell_key: str) -> tuple[NodeKey, str, str, int]:
    """Normalize `cell_key` and return `(canonical, sheet, column, row)`.

    Raises:
        ValueError: If `cell_key` is a range / row key rather than a single cell.
    """
    try:
        parsed = _parse_node_key(cell_key)
    except ValueError as exc:
        raise ValueError(f"Expected a single-cell key, got: {cell_key!r}") from exc
    if not isinstance(parsed, CellKey):
        raise ValueError(f"Expected a single-cell key, got range/union key: {cell_key!r}")
    return str(parsed), parsed.sheet, parsed.column, parsed.row


def _require_one_row_key(range_key: str) -> tuple[NodeKey, ParsedRowKey]:
    """Canonicalize a one-row range key and return `(canonical, parsed)`."""
    canonical = _normalize_row_key(range_key)
    return canonical, _parse_row_key(canonical)


def _row_node_span_indices(
    row_node: Node | NodeView,
) -> tuple[str, int, int, int]:
    """Return `(sheet, row, min_col_idx, max_col_idx)` for a row node."""
    if row_node.kind is not NodeKind.row and row_node.shape is not NodeShape.row:
        raise ValueError("Expected a row node")
    if (
        row_node.sheet is None
        or row_node.min_col is None
        or row_node.max_col is None
        or row_node.row is None
    ):
        raise ValueError("Row node is missing extent fields")
    lo = fastpyxl.utils.cell.column_index_from_string(row_node.min_col)
    hi = fastpyxl.utils.cell.column_index_from_string(row_node.max_col)
    return row_node.sheet, row_node.row, lo, hi


def _span_contains(
    *,
    owner_sheet: str,
    owner_row: int,
    owner_lo: int,
    owner_hi: int,
    sheet: str,
    row: int,
    min_col: str,
    max_col: str,
) -> bool:
    if sheet != owner_sheet or row != owner_row:
        return False
    q_lo = fastpyxl.utils.cell.column_index_from_string(min_col)
    q_hi = fastpyxl.utils.cell.column_index_from_string(max_col)
    if q_lo > q_hi:
        q_lo, q_hi = q_hi, q_lo
    return owner_lo <= q_lo and q_hi <= owner_hi


def row_node_covers_cell(row_node: Node | NodeView, cell_key: str) -> bool:
    """Return True if `cell_key` lies in `row_node`'s one-row column span.

    Raises:
        ValueError: If `row_node` is not a row node, or `cell_key` is not a
            single-cell address.
    """
    sheet_o, row_o, lo, hi = _row_node_span_indices(row_node)
    _cell_key, sheet, col, row = _require_single_cell_key(cell_key)
    return _span_contains(
        owner_sheet=sheet_o,
        owner_row=row_o,
        owner_lo=lo,
        owner_hi=hi,
        sheet=sheet,
        row=row,
        min_col=col,
        max_col=col,
    )


def row_node_covers_range(row_node: Node | NodeView, range_key: str) -> bool:
    """Return True if one-row `range_key` is fully contained in `row_node`.

    The query is canonicalized first (column order, `$` stripped, both-end sheet
    forms collapsed). Partial overlaps return False.

    Raises:
        ValueError: If `row_node` is not a row node, or `range_key` is not a
            valid one-row range.
    """
    sheet_o, row_o, lo, hi = _row_node_span_indices(row_node)
    _canonical, parsed = _require_one_row_key(range_key)
    return _span_contains(
        owner_sheet=sheet_o,
        owner_row=row_o,
        owner_lo=lo,
        owner_hi=hi,
        sheet=parsed.sheet,
        row=parsed.row,
        min_col=parsed.min_col,
        max_col=parsed.max_col,
    )


class _GraphNodeLookup(Protocol):
    def __iter__(self) -> Iterator[NodeKey]: ...

    def get_node(self, key: NodeKey) -> NodeView | None: ...


def find_nodes_covering_cell(graph: _GraphNodeLookup, cell_key: str) -> list[NodeKey]:
    """Return multi-cell node keys that own `cell_key`."""
    canonical, _sheet, _col, _row = _require_single_cell_key(cell_key)
    out: list[NodeKey] = []
    for key in graph:
        node = graph.get_node(key)
        if node is None or node.kind is NodeKind.cell or node.address is None:
            continue
        if canonical in member_keys(node):
            out.append(key)
    return out


def _one_row_query_cells(sheet: str, row: int, min_col: str, max_col: str) -> list[NodeKey]:
    lo = fastpyxl.utils.cell.column_index_from_string(min_col)
    hi = fastpyxl.utils.cell.column_index_from_string(max_col)
    if lo > hi:
        lo, hi = hi, lo
    return [
        _format_cell_key(sheet, fastpyxl.utils.cell.get_column_letter(c), row)
        for c in range(lo, hi + 1)
    ]


def _multi_cell_covers_one_row(
    node: Node | NodeView, sheet: str, row: int, min_col: str, max_col: str
) -> bool:
    """Return True if every cell in the one-row query is a member of `node`."""
    if node.kind is NodeKind.cell or node.address is None:
        return False
    if node.kind is NodeKind.row:
        sheet_o, row_o, lo, hi = _row_node_span_indices(node)
        return _span_contains(
            owner_sheet=sheet_o,
            owner_row=row_o,
            owner_lo=lo,
            owner_hi=hi,
            sheet=sheet,
            row=row,
            min_col=min_col,
            max_col=max_col,
        )
    members = set(member_keys(node))
    return all(c in members for c in _one_row_query_cells(sheet, row, min_col, max_col))


def find_row_nodes_covering(graph: _GraphNodeLookup, address: str) -> list[NodeKey]:
    """Return multi-cell node keys that fully cover a cell or one-row subrange.

    Includes one-row `RangeKey` nodes (`kind=row`) and any `union`/`range`
    owner whose expanded members contain every cell of the query.

    `address` is canonicalized first. For a one-row range, coverage means the
    owner contains the query span. For a cell, coverage means the cell lies in
    the owner.
    """
    try:
        _canonical, parsed = _require_one_row_key(address)
        sheet, row, min_col, max_col = (
            parsed.sheet,
            parsed.row,
            parsed.min_col,
            parsed.max_col,
        )
    except ValueError:
        _cell_key, sheet, col, row = _require_single_cell_key(address)
        min_col = max_col = col

    out: list[NodeKey] = []
    for key in graph:
        node = graph.get_node(key)
        if node is None or node.kind is NodeKind.cell:
            continue
        if _multi_cell_covers_one_row(node, sheet, row, min_col, max_col):
            out.append(key)
    return out


def locate_range(graph: _GraphNodeLookup, range_key: str) -> RangeLocation | None:
    """Return where a one-row range is represented in `graph`.

    The query is canonicalized via `normalize_row_key`. Resolution order:
    1. Exact multi-cell node at the canonical key (`row` stripe or other).
    2. Otherwise the unique multi-cell node whose members fully contain the query
       (row span or union/range membership).

    Returns:
        `RangeLocation` when represented, else `None`.

    Raises:
        ValueError: If `range_key` is not a one-row range, or unique occupancy is
            violated (multiple covering nodes).
    """
    canonical, parsed = _require_one_row_key(range_key)
    exact = graph.get_node(canonical)
    if exact is not None and exact.kind is not NodeKind.cell:
        return RangeLocation(
            range_key=canonical,
            kind=exact.kind,
            node_key=canonical,
            min_col=parsed.min_col,
            max_col=parsed.max_col,
            row=parsed.row,
        )

    covering = find_row_nodes_covering(graph, canonical)
    if len(covering) > 1:
        keys = ", ".join(covering)
        raise ValueError(
            f"Unique cell occupancy violated for {canonical}: covered by overlapping nodes {keys}"
        )
    if len(covering) == 1:
        owner = graph.get_node(covering[0])
        assert owner is not None
        return RangeLocation(
            range_key=canonical,
            kind=owner.kind,
            node_key=covering[0],
            min_col=parsed.min_col,
            max_col=parsed.max_col,
            row=parsed.row,
        )
    return None


def locate_cell(graph: _GraphNodeLookup, cell_key: str) -> CellLocation | None:
    """Return where a single cell is represented in `graph`.

    Prefer this over `DependencyGraph.get_node` when resolving a workbook cell
    that may be a member of a multi-cell (`RangeKey` / `UnionKey`) node.
    `get_node` is exact-key only and returns `None` for members.

    Resolution order:
    1. Occupancy index via `graph.cell_owner` when available (O(1)).
    2. Exact `kind=cell` node at `cell_key`.
    3. Otherwise the unique covering multi-cell node (`row` or `union`), if any.

    Returns:
        `CellLocation` when the cell is represented, else `None`.

    Raises:
        ValueError: If `cell_key` is not a single cell, or unique cell occupancy
            is violated (cell node plus covering multi-cell, or multiple covers).
    """
    canonical, _sheet, column, _row = _require_single_cell_key(cell_key)

    cell_owner_fn = getattr(graph, "cell_owner", None)
    if callable(cell_owner_fn):
        owner_key = cell_owner_fn(canonical)
        if owner_key is None:
            return None
        owner = graph.get_node(owner_key)
        if owner is None:
            return None
        cell_node = graph.get_node(canonical)
        if cell_node is not None and cell_node.kind is NodeKind.cell and owner_key != canonical:
            raise ValueError(
                f"Unique cell occupancy violated for {canonical}: "
                f"cell node exists and is also covered by node(s) {owner_key}"
            )
        return CellLocation(
            cell_key=canonical,
            kind=owner.kind,
            node_key=owner_key,
            column=column,
        )

    cell_node = graph.get_node(canonical)
    covering = find_nodes_covering_cell(graph, canonical)

    if cell_node is not None and cell_node.kind is NodeKind.cell:
        if covering:
            keys = ", ".join(covering)
            raise ValueError(
                f"Unique cell occupancy violated for {canonical}: "
                f"cell node exists and is also covered by node(s) {keys}"
            )
        return CellLocation(
            cell_key=canonical,
            kind=NodeKind.cell,
            node_key=canonical,
            column=column,
        )

    if len(covering) > 1:
        keys = ", ".join(covering)
        raise ValueError(
            f"Unique cell occupancy violated for {canonical}: covered by overlapping nodes {keys}"
        )
    if len(covering) == 1:
        owner = graph.get_node(covering[0])
        assert owner is not None
        return CellLocation(
            cell_key=canonical,
            kind=owner.kind,
            node_key=covering[0],
            column=column,
        )
    return None
