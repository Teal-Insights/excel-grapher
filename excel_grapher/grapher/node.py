from __future__ import annotations

from collections import OrderedDict
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from enum import StrEnum
from types import MappingProxyType
from typing import Any, Protocol, TypeAlias

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    CellKey,
    NodeShape,
    RangeKey,
    UnionKey,
)
from excel_grapher.core.address_keys import NodeKey as AddressKey
from excel_grapher.core.address_keys import expand_node_cells as _expand_node_cells
from excel_grapher.core.address_keys import format_cell_key as _format_cell_key
from excel_grapher.core.address_keys import members_to_node_key as _members_to_node_key
from excel_grapher.core.address_keys import parse_address as _parse_address
from excel_grapher.core.address_keys import parse_node_key as _parse_node_key
from excel_grapher.core.address_keys import quote_sheet_if_needed as _quote_sheet

# Graph node identity; same canonical form as NormalizedAddress (cell / range / union).
NodeKey: TypeAlias = str

# Bounded cache for Node derived fields (`key` / `shape` / `column_index`).
# External to the instance so `@dataclass(slots=True)` stays compatible (unlike
# `functools.cached_property`, which needs a per-instance `__dict__`).
# Use a dict LRU (not `functools.lru_cache`): `_make_key` treats `str` subclasses
# as distinct from plain `str`, which would split AddressKey / str traffic.
_NODE_DERIVED_CACHE_MAXSIZE = 16384


@dataclass(frozen=True, slots=True)
class _NodeDerivedFields:
    """Cached derived Node attributes for one canonical address."""

    key: NodeKey
    shape: NodeShape
    column_index: int | None


@dataclass(frozen=True, slots=True)
class NodeDerivedCacheInfo:
    """Hit/miss statistics for the Node derived-fields LRU."""

    hits: int
    misses: int
    maxsize: int
    currsize: int


def _compute_derived_fields(address: AddressKey | str) -> _NodeDerivedFields:
    """Compute derived Node fields for a canonical address."""
    if isinstance(address, (CellKey, RangeKey, UnionKey)):
        parsed: AddressKey = address
    else:
        parsed = _parse_node_key(address)
    key = address if type(address) is str else str(address)
    column_index: int | None = None
    if isinstance(parsed, CellKey):
        column_index = int(fastpyxl.utils.cell.column_index_from_string(parsed.column))
    return _NodeDerivedFields(key=key, shape=parsed.shape, column_index=column_index)


class _NodeDerivedFieldsCache:
    """Process-wide LRU for Node derived fields, keyed by address text."""

    def __init__(self, maxsize: int = _NODE_DERIVED_CACHE_MAXSIZE) -> None:
        if maxsize < 1:
            raise ValueError("maxsize must be at least 1")
        self._maxsize = maxsize
        self._cache: OrderedDict[str, _NodeDerivedFields] = OrderedDict()
        self._hits = 0
        self._misses = 0

    def get(self, address: AddressKey | str) -> _NodeDerivedFields:
        cached = self._cache.get(address)
        if cached is not None:
            self._hits += 1
            self._cache.move_to_end(address)
            return cached

        self._misses += 1
        derived = _compute_derived_fields(address)
        # Store under a plain str so AddressKey and str lookups share one entry.
        store_key = address if type(address) is str else str(address)
        self._cache[store_key] = derived
        if len(self._cache) > self._maxsize:
            self._cache.popitem(last=False)
        return derived

    def clear(self) -> None:
        self._cache.clear()
        self._hits = 0
        self._misses = 0

    def cache_info(self) -> NodeDerivedCacheInfo:
        return NodeDerivedCacheInfo(
            hits=self._hits,
            misses=self._misses,
            maxsize=self._maxsize,
            currsize=len(self._cache),
        )


_DERIVED_FIELDS_CACHE = _NodeDerivedFieldsCache()


def _lookup_derived_fields(address: AddressKey | str) -> _NodeDerivedFields:
    """Return cached derived fields for `address` (plain str or AddressKey)."""
    return _DERIVED_FIELDS_CACHE.get(address)


def _derived_fields(address: AddressKey) -> _NodeDerivedFields:
    return _lookup_derived_fields(address)


def _derived_fields_cache_info() -> NodeDerivedCacheInfo:
    """Return hit/miss statistics for the derived-fields LRU."""
    return _DERIVED_FIELDS_CACHE.cache_info()


def _derived_fields_cache_clear() -> None:
    """Clear the derived-fields LRU (intended for tests)."""
    _DERIVED_FIELDS_CACHE.clear()


class _EmptyMetadata(Mapping[str, Any]):
    """Immutable empty mapping shared by every node with no metadata.

    Nodes default to the `EMPTY_METADATA` singleton rather than a fresh dict, so
    the common (empty) case costs one shared object instead of ~64 bytes per
    node. Reads behave like `{}`; writes raise `TypeError` because the mapping is
    not mutable — populate metadata via `Node.set_metadata` /
    `Node.update_metadata` (or `DependencyGraph.set_node_metadata`).
    """

    __slots__ = ()

    def __getitem__(self, key: str) -> Any:
        raise KeyError(key)

    def __iter__(self) -> Iterator[str]:
        return iter(())

    def __len__(self) -> int:
        return 0

    def __contains__(self, key: object) -> bool:
        return False

    def __hash__(self) -> int:
        # `Mapping` sets `__hash__ = None`; instances are immutable and all equal,
        # so a constant hash is consistent. Also required for a dataclass default.
        return hash(())

    def __repr__(self) -> str:
        return "{}"

    def __copy__(self) -> _EmptyMetadata:
        return self

    def __deepcopy__(self, memo: dict[int, Any]) -> _EmptyMetadata:
        return self

    def __reduce__(self) -> tuple[Any, ...]:
        return (_empty_metadata, ())


EMPTY_METADATA: Mapping[str, Any] = _EmptyMetadata()


def _empty_metadata() -> Mapping[str, Any]:
    """Return the shared empty-metadata singleton (pickle/copy reconstructor)."""
    return EMPTY_METADATA


def copy_metadata(metadata: Mapping[str, Any] | None) -> Mapping[str, Any]:
    """Return an independent copy of `metadata`, sharing the empty singleton.

    Empty (or missing) metadata propagates as `EMPTY_METADATA` without
    materializing a dict; anything else is copied into a plain dict.
    """
    return dict(metadata) if metadata else EMPTY_METADATA


class NodeKind(StrEnum):
    """Coarse node classification for compatibility metadata.

    Prefer `Node.shape` / `parse_node_key(node.address)` for geometry. Multi-cell
    addresses (row/column/range/union shapes) all use `union`.
    """

    cell = "cell"
    union = "union"


def _kind_for_address(address: AddressKey) -> NodeKind:
    if isinstance(address, CellKey):
        return NodeKind.cell
    return NodeKind.union


def _coerce_address(value: str | AddressKey) -> AddressKey:
    """Canonicalize `value` via `parse_node_key` (1x1 ranges collapse to cells)."""
    return _parse_node_key(str(value))


@dataclass(slots=True)
class Node:
    """Workbook cell or multi-cell span in a dependency graph.

    Canonical identity is `address` (`CellKey` / `RangeKey` / `UnionKey`). Legacy
    constructors that pass `sheet`/`column`/`row` still work and sync `address`
    in `__post_init__`. Contiguous rectangles may also be built from sheet +
    extent fields (`min_col`/`max_col`/`min_row`/`max_row`) without a scalar
    `column`.

    Multi-cell nodes store `formula` / `normalized_formula` / `value` as `None`.

    `is_leaf` is true when the node has no outgoing dependency edges (value-only
    cells and literal-only formulas such as `=1+1`).

    Instances are slotted (no per-instance `__dict__`). Derived attributes
    (`key`, `shape`, `column_index`) are served from a process-wide LRU keyed on
    the canonical address rather than `functools.cached_property`.

    `metadata` is a read-only `Mapping` for callers. Nodes without metadata share
    the immutable `EMPTY_METADATA` singleton instead of allocating a dict, so
    `node.metadata["k"] = v` is not supported (it raises `TypeError` on a node
    with no metadata yet). Writers — node hooks included — use
    `set_metadata` / `update_metadata`, or `DependencyGraph.set_node_metadata`.
    """

    sheet: str | None
    column: str | None
    row: int | None
    formula: str | None
    normalized_formula: str | None
    value: Any
    is_leaf: bool
    is_target: bool = False
    metadata: Mapping[str, Any] = EMPTY_METADATA
    kind: NodeKind = NodeKind.cell
    min_col: str | None = None
    min_row: int | None = None
    max_col: str | None = None
    max_row: int | None = None
    address: AddressKey | None = None

    def __post_init__(self) -> None:
        if not self.metadata:
            # Never hold a caller's empty dict; share the singleton instead.
            self.metadata = EMPTY_METADATA

        if self.address is not None:
            self.address = _coerce_address(self.address)
            self._sync_fields_from_address()
            return

        # Extent-only construction for contiguous rectangles (no scalar column).
        if (
            self.min_col is not None
            and self.max_col is not None
            and self.min_row is not None
            and self.max_row is not None
            and self.column is None
            and self.sheet is not None
        ):
            lo_row, hi_row = sorted((int(self.min_row), int(self.max_row)))
            left = str(self.min_col).upper()
            right = str(self.max_col).upper()
            if fastpyxl.utils.cell.column_index_from_string(
                left
            ) > fastpyxl.utils.cell.column_index_from_string(right):
                left, right = right, left
            text = f"{_quote_sheet(self.sheet)}!{left}{lo_row}:{right}{hi_row}"
            self.address = _coerce_address(text)
            self._sync_fields_from_address()
            return

        self._finalize_cell_legacy()
        assert self.sheet is not None and self.column is not None and self.row is not None
        self.address = CellKey(_format_cell_key(self.sheet, self.column, self.row))
        self.kind = NodeKind.cell

    def _clear_multi_cell_formula_fields(self) -> None:
        if self.formula is not None or self.normalized_formula is not None:
            raise ValueError(
                "Multi-cell nodes cannot have formula or normalized_formula; leave both None"
            )
        self.formula = None
        self.normalized_formula = None
        self.value = None

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

        self._clear_multi_cell_formula_fields()
        if isinstance(addr, RangeKey):
            self.sheet = addr.sheet
            self.min_col = addr.min_col
            self.max_col = addr.max_col
            self.min_row = addr.min_row
            self.max_row = addr.max_row
            self.column = addr.column
            self.row = addr.row
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

    def set_metadata(self, metadata: Mapping[str, Any] | None) -> None:
        """Replace this node's metadata with a copy of `metadata`.

        Empty or `None` input drops back to the shared `EMPTY_METADATA`
        singleton, so nodes without metadata never own a dict.
        """
        self.metadata = copy_metadata(metadata)

    def update_metadata(self, updates: Mapping[str, Any]) -> None:
        """Merge `updates` into this node's metadata, allocating on first write."""
        if not updates:
            return
        current = self.metadata
        if isinstance(current, dict):
            current.update(updates)
        else:
            self.metadata = {**current, **updates}

    @property
    def key(self) -> NodeKey:
        assert self.address is not None
        return _derived_fields(self.address).key

    @property
    def shape(self) -> NodeShape:
        assert self.address is not None
        return _derived_fields(self.address).shape

    @property
    def column_index(self) -> int:
        assert self.address is not None
        idx = _derived_fields(self.address).column_index
        if idx is None:
            raise ValueError("column_index is only defined for cell nodes")
        return idx


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

    @property
    def key(self) -> NodeKey:
        if self.address is not None:
            return str(self.address)
        assert self.sheet is not None and self.column is not None and self.row is not None
        return _format_cell_key(self.sheet, self.column, self.row)

    @property
    def shape(self) -> NodeShape:
        if self.address is not None:
            return self.address.shape
        if self.kind is NodeKind.union:
            return NodeShape.union
        return NodeShape.cell

    @property
    def column_index(self) -> int:
        if self.kind is not NodeKind.cell or self.column is None:
            raise ValueError("column_index is only defined for cell nodes")
        return int(fastpyxl.utils.cell.column_index_from_string(self.column))


def _view_metadata(metadata: Mapping[str, Any]) -> Mapping[str, Any]:
    """Return a read-only snapshot of `metadata` for a `NodeView`."""
    return MappingProxyType(dict(metadata)) if metadata else EMPTY_METADATA


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
        metadata=_view_metadata(node.metadata),
        kind=node.kind,
        min_col=node.min_col,
        min_row=node.min_row,
        max_col=node.max_col,
        max_row=node.max_row,
        address=node.address,
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
    metadata: Mapping[str, Any] | None = None,
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
        metadata=copy_metadata(metadata),
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
    metadata: Mapping[str, Any] | None = None,
) -> Node:
    """Build a multi-cell node from cell members (or a cell node for one member).

    Multi-cell nodes require `formula` / `normalized_formula` / `value` to be
    `None`. A single member collapses to a cell node via `members_to_node_key`.
    """
    if not members:
        raise ValueError("Cannot build node from empty member set")
    address = _members_to_node_key(members)
    if isinstance(address, CellKey):
        return Node(
            sheet=None,
            column=None,
            row=None,
            formula=formula,
            normalized_formula=normalized_formula,
            value=value,
            is_leaf=is_leaf,
            is_target=is_target,
            metadata=copy_metadata(metadata),
            address=address,
        )
    if formula is not None or normalized_formula is not None:
        raise ValueError(
            "Multi-cell nodes cannot have formula or normalized_formula; leave both None"
        )
    return Node(
        sheet=None,
        column=None,
        row=None,
        formula=None,
        normalized_formula=None,
        value=None,
        is_leaf=is_leaf,
        is_target=is_target,
        metadata=copy_metadata(metadata),
        address=address,
    )


def member_keys(node: Node | NodeView) -> list[NodeKey]:
    """Return sheet-qualified cell keys owned by `node` (expansion of `address`)."""
    if node.address is None:
        raise ValueError("Node is missing address")
    return [str(c) for c in _expand_node_cells(node.address)]


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
        kind: Owning node kind (`cell` or `union`).
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
        kind: Owning node kind (`union` for multi-cell owners).
        node_key: Graph key of the owning multi-cell node.
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
        ValueError: If `cell_key` is a range / union key rather than a single cell.
    """
    try:
        parsed = _parse_node_key(cell_key)
    except ValueError as exc:
        raise ValueError(f"Expected a single-cell key, got: {cell_key!r}") from exc
    if not isinstance(parsed, CellKey):
        raise ValueError(f"Expected a single-cell key, got range/union key: {cell_key!r}")
    return str(parsed), parsed.sheet, parsed.column, parsed.row


def _require_one_row_range(range_key: str) -> RangeKey:
    """Canonicalize a one-row range key (multi-column; 1x1 collapses to a cell)."""
    try:
        parsed = _parse_node_key(range_key)
    except ValueError as exc:
        raise ValueError(f"Expected a one-row range key, got: {range_key!r}") from exc
    if not isinstance(parsed, RangeKey) or parsed.shape is not NodeShape.row:
        raise ValueError(f"Expected a one-row range key, got: {range_key!r}")
    return parsed


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
    if node.shape is NodeShape.row and node.sheet == sheet and node.row == row:
        if node.min_col is None or node.max_col is None:
            return False
        lo = fastpyxl.utils.cell.column_index_from_string(node.min_col)
        hi = fastpyxl.utils.cell.column_index_from_string(node.max_col)
        q_lo = fastpyxl.utils.cell.column_index_from_string(min_col)
        q_hi = fastpyxl.utils.cell.column_index_from_string(max_col)
        if q_lo > q_hi:
            q_lo, q_hi = q_hi, q_lo
        return lo <= q_lo and q_hi <= hi
    members = set(member_keys(node))
    return all(c in members for c in _one_row_query_cells(sheet, row, min_col, max_col))


def _find_nodes_covering_one_row(
    graph: _GraphNodeLookup, sheet: str, row: int, min_col: str, max_col: str
) -> list[NodeKey]:
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

    Resolution order:
    1. Exact multi-cell node at the canonical key.
    2. Otherwise the unique multi-cell node whose members fully contain the query.

    Returns:
        `RangeLocation` when represented, else `None`.

    Raises:
        ValueError: If `range_key` is not a one-row range, or unique occupancy is
            violated (multiple covering nodes).
    """
    parsed = _require_one_row_range(range_key)
    canonical = str(parsed)
    exact = graph.get_node(canonical)
    if exact is not None and exact.kind is not NodeKind.cell:
        return RangeLocation(
            range_key=canonical,
            kind=exact.kind,
            node_key=canonical,
            min_col=parsed.min_col,
            max_col=parsed.max_col,
            row=parsed.min_row,
        )

    covering = _find_nodes_covering_one_row(
        graph, parsed.sheet, parsed.min_row, parsed.min_col, parsed.max_col
    )
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
            row=parsed.min_row,
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
    3. Otherwise the unique covering multi-cell node, if any.

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
