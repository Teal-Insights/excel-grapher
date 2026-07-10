from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from enum import StrEnum
from types import MappingProxyType
from typing import Any, Protocol, TypeAlias

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import NormalizedAddress
from excel_grapher.core.address_keys import format_cell_key as _format_cell_key
from excel_grapher.core.address_keys import format_row_key as _format_row_key
from excel_grapher.core.address_keys import normalize_key as _normalize_key
from excel_grapher.core.address_keys import parse_address as _parse_address
from excel_grapher.core.address_keys import parse_row_key as _parse_row_key

# Graph node identity; same canonical form as NormalizedAddress (cell or row key).
NodeKey: TypeAlias = NormalizedAddress


class NodeKind(StrEnum):
    """Kind of dependency-graph node."""

    cell = "cell"
    row = "row"


@dataclass
class Node:
    """Workbook cell or one-row span in a dependency graph.

    Cell nodes (`kind=cell`) keep the historical single-cell shape. Row nodes
    (`kind=row`) represent a horizontal stripe on one worksheet row with column
    span `min_col`…`max_col` (`min_row == max_row == row`).

    `is_leaf` is true when the node has no outgoing dependency edges (value-only
    cells and literal-only formulas such as `=1+1`).
    """

    sheet: str
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

    def __post_init__(self) -> None:
        if self.kind is NodeKind.cell:
            self._finalize_cell()
        elif self.kind is NodeKind.row:
            self._finalize_row()
        else:
            raise ValueError(f"Unsupported node kind: {self.kind!r}")

    def _finalize_cell(self) -> None:
        if self.column is None or self.row is None:
            raise ValueError("Cell nodes require column and row")
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

    def _finalize_row(self) -> None:
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

    @property
    def key(self) -> NodeKey:
        if self.kind is NodeKind.row:
            assert self.min_col is not None and self.max_col is not None and self.row is not None
            return _format_row_key(self.sheet, self.min_col, self.row, self.max_col)
        assert self.column is not None and self.row is not None
        return _format_cell_key(self.sheet, self.column, self.row)

    @property
    def address(self) -> str:
        if self.kind is NodeKind.row:
            assert self.min_col is not None and self.max_col is not None and self.row is not None
            return f"{self.min_col}{self.row}:{self.max_col}{self.row}"
        assert self.column is not None and self.row is not None
        return f"{self.column}{self.row}"

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

    sheet: str
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

    @property
    def key(self) -> NodeKey:
        if self.kind is NodeKind.row:
            assert self.min_col is not None and self.max_col is not None and self.row is not None
            return _format_row_key(self.sheet, self.min_col, self.row, self.max_col)
        assert self.column is not None and self.row is not None
        return _format_cell_key(self.sheet, self.column, self.row)

    @property
    def address(self) -> str:
        if self.kind is NodeKind.row:
            assert self.min_col is not None and self.max_col is not None and self.row is not None
            return f"{self.min_col}{self.row}:{self.max_col}{self.row}"
        assert self.column is not None and self.row is not None
        return f"{self.column}{self.row}"

    @property
    def column_index(self) -> int:
        if self.kind is not NodeKind.cell or self.column is None:
            raise ValueError("column_index is only defined for cell nodes")
        return int(fastpyxl.utils.cell.column_index_from_string(self.column))


def node_to_view(node: Node) -> NodeView:
    """Build an immutable `NodeView` snapshot from a stored `Node`."""
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

    Args:
        sheet: Worksheet name.
        row: Worksheet row index (1-based); both endpoints use this row.
        min_col: Left column letters (ordered with `max_col` on construction).
        max_col: Right column letters.
        formula: Optional formula text stored on the node.
        normalized_formula: Optional sheet-qualified formula text.
        value: Optional cached value.
        is_leaf: Whether the node has no outgoing dependency edges.
        is_target: Whether the node is a graph target.
        metadata: Optional metadata dict (copied).

    Returns:
        A `Node` with `kind=row` whose key is the canonical one-row span.

    Raises:
        ValueError: If the resulting extent is not a single worksheet row.
    """
    return Node(
        sheet=sheet,
        column=None,
        row=row,
        formula=formula,
        normalized_formula=normalized_formula,
        value=value,
        is_leaf=is_leaf,
        is_target=is_target,
        metadata=dict(metadata or {}),
        kind=NodeKind.row,
        min_col=min_col,
        min_row=row,
        max_col=max_col,
        max_row=row,
    )


def row_member_keys(node: Node | NodeView) -> list[NodeKey]:
    """Return sheet-qualified cell keys for each column in a row node's span.

    Members share sheet and row; only the column differs. Raises `ValueError`
    when `node` is not a row node.
    """
    if node.kind is not NodeKind.row:
        raise ValueError("row_member_keys requires a row node")
    if node.min_col is None or node.max_col is None or node.row is None:
        raise ValueError("Row node is missing extent fields")
    start = fastpyxl.utils.cell.column_index_from_string(node.min_col)
    end = fastpyxl.utils.cell.column_index_from_string(node.max_col)
    keys: list[NodeKey] = []
    for col_idx in range(start, end + 1):
        col = fastpyxl.utils.cell.get_column_letter(col_idx)
        keys.append(_format_cell_key(node.sheet, col, node.row))
    return keys


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
        kind: `cell` when the graph has a cell node; `row` when the cell lives
            only inside a row node's span.
        node_key: Graph key of the owning node (`Sheet1!E63` or `Sheet1!D63:Y63`).
        column: Column letters of the queried cell.
    """

    cell_key: NodeKey
    kind: NodeKind
    node_key: NodeKey
    column: str


def _require_single_cell_key(cell_key: str) -> tuple[NodeKey, str, str, int]:
    """Normalize `cell_key` and return `(canonical, sheet, column, row)`.

    Raises:
        ValueError: If `cell_key` is a range / row key rather than a single cell.
    """
    try:
        _parse_row_key(cell_key)
    except ValueError:
        pass
    else:
        raise ValueError(f"Expected a single-cell key, got row/range key: {cell_key!r}")

    try:
        normalized = _normalize_key(cell_key)
        sheet, coord = _parse_address(normalized)
        col, row = fastpyxl.utils.cell.coordinate_from_string(coord.replace("$", ""))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Expected a single-cell key, got: {cell_key!r}") from exc
    col_u = str(col).upper()
    row_i = int(row)
    return _format_cell_key(sheet, col_u, row_i), sheet, col_u, row_i


def row_node_covers_cell(row_node: Node | NodeView, cell_key: str) -> bool:
    """Return True if `cell_key` lies in `row_node`'s one-row column span.

    Raises:
        ValueError: If `row_node` is not a row node, or `cell_key` is not a
            single-cell address.
    """
    if row_node.kind is not NodeKind.row:
        raise ValueError("row_node_covers_cell requires a row node")
    if row_node.min_col is None or row_node.max_col is None or row_node.row is None:
        raise ValueError("Row node is missing extent fields")

    _cell_key, sheet, col, row = _require_single_cell_key(cell_key)
    if sheet != row_node.sheet or row != row_node.row:
        return False
    c = fastpyxl.utils.cell.column_index_from_string(col)
    lo = fastpyxl.utils.cell.column_index_from_string(row_node.min_col)
    hi = fastpyxl.utils.cell.column_index_from_string(row_node.max_col)
    return lo <= c <= hi


class _GraphNodeLookup(Protocol):
    def __iter__(self) -> Iterator[NodeKey]: ...

    def get_node(self, key: NodeKey) -> NodeView | None: ...


def find_row_nodes_covering(graph: _GraphNodeLookup, cell_key: str) -> list[NodeKey]:
    """Return row-node keys whose span includes `cell_key` (insertion order)."""
    _require_single_cell_key(cell_key)
    out: list[NodeKey] = []
    for key in graph:
        node = graph.get_node(key)
        if node is None or node.kind is not NodeKind.row:
            continue
        if row_node_covers_cell(node, cell_key):
            out.append(key)
    return out


def locate_cell(graph: _GraphNodeLookup, cell_key: str) -> CellLocation | None:
    """Return where a single cell is represented in `graph`.

    Resolution order:
    1. Exact `kind=cell` node at `cell_key`.
    2. Otherwise the unique covering row node, if any.

    Returns:
        `CellLocation` when the cell is represented, else `None`.

    Raises:
        ValueError: If `cell_key` is not a single cell, or unique cell occupancy
            is violated (cell node plus covering row, or multiple covering rows).
    """
    canonical, _sheet, column, _row = _require_single_cell_key(cell_key)
    cell_node = graph.get_node(canonical)
    covering = find_row_nodes_covering(graph, canonical)

    if cell_node is not None and cell_node.kind is NodeKind.cell:
        if covering:
            keys = ", ".join(covering)
            raise ValueError(
                f"Unique cell occupancy violated for {canonical}: "
                f"cell node exists and is also covered by row node(s) {keys}"
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
            f"Unique cell occupancy violated for {canonical}: "
            f"covered by overlapping row nodes {keys}"
        )
    if len(covering) == 1:
        return CellLocation(
            cell_key=canonical,
            kind=NodeKind.row,
            node_key=covering[0],
            column=column,
        )
    return None
