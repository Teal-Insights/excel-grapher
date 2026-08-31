from __future__ import annotations

from collections import OrderedDict
from collections.abc import Iterator, Mapping
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import Any, TypeAlias

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    CellKey,
    NodeShape,
)
from excel_grapher.core.address_keys import format_cell_key as _format_cell_key
from excel_grapher.core.address_keys import parse_node_key as _parse_node_key
from excel_grapher.core.formula_ast import (
    AstNode,
    FormulaStyle,
    parse_formula_text,
    render_formula,
)

# Graph node identity; canonical sheet-qualified cell address string.
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
    column_index: int


@dataclass(frozen=True, slots=True)
class NodeDerivedCacheInfo:
    """Hit/miss statistics for the Node derived-fields LRU."""

    hits: int
    misses: int
    maxsize: int
    currsize: int


def _coerce_cell_address(value: str | CellKey) -> CellKey:
    """Canonicalize `value` to a `CellKey` (1x1 ranges collapse to cells)."""
    parsed = _parse_node_key(str(value))
    if not isinstance(parsed, CellKey):
        raise ValueError(
            f"Graph nodes must be single cells; got {type(parsed).__name__}: {value!r}"
        )
    return parsed


def _compute_derived_fields(address: CellKey | str) -> _NodeDerivedFields:
    """Compute derived Node fields for a canonical cell address."""
    parsed = address if isinstance(address, CellKey) else _coerce_cell_address(address)
    key = address if type(address) is str else str(address)
    column_index = int(fastpyxl.utils.cell.column_index_from_string(parsed.column))
    return _NodeDerivedFields(key=key, shape=NodeShape.cell, column_index=column_index)


class _NodeDerivedFieldsCache:
    """Process-wide LRU for Node derived fields, keyed by address text."""

    def __init__(self, maxsize: int = _NODE_DERIVED_CACHE_MAXSIZE) -> None:
        if maxsize < 1:
            raise ValueError("maxsize must be at least 1")
        self._maxsize = maxsize
        self._cache: OrderedDict[str, _NodeDerivedFields] = OrderedDict()
        self._hits = 0
        self._misses = 0

    def get(self, address: CellKey | str) -> _NodeDerivedFields:
        cached = self._cache.get(address)
        if cached is not None:
            self._hits += 1
            self._cache.move_to_end(address)
            return cached

        self._misses += 1
        derived = _compute_derived_fields(address)
        # Store under a plain str so CellKey and str lookups share one entry.
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


def _lookup_derived_fields(address: CellKey | str) -> _NodeDerivedFields:
    """Return cached derived fields for `address` (plain str or CellKey)."""
    return _DERIVED_FIELDS_CACHE.get(address)


def _derived_fields(address: CellKey) -> _NodeDerivedFields:
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


@dataclass(slots=True, init=False)
class Node:
    """Single workbook cell in a dependency graph.

    Canonical identity is `address` (`CellKey`). Change it through
    `DependencyGraph.move_node`; assigning `address` directly raises so relative
    formula axes cannot silently retarget. Legacy constructors that pass
    `sheet`/`column`/`row` still work and sync `address` in `__init__`.

    `formula_ast` is the primary in-memory formula artifact. Construct nodes
    with `formula_ast=` when the tree is already known. `normalized_formula` is
    a derived view (`render_formula` with `A1_ABSOLUTE`); it is not stored, and
    the constructor keyword of the same name is only a bootstrap for string
    input. When an address is available, that bootstrap parses with
    `parse_preserving_axes` so `$` vs bare A1 is kept; without an anchor,
    relative intent cannot be recovered. Unparseable cells keep the last known
    formula text as a fallback so `has_formula` still holds.
    `DependencyGraph.formula_shapes` is an optional eval/codegen overlay;
    missing or dropped shapes fall back to this AST. The raw workbook string
    `formula` is opt-in at extraction
    (`create_dependency_graph(store_raw_formula=True)`) and audit-only -- on a
    compressed graph it still holds the pre-compression text, so never re-parse
    it as the node's current definition.

    `is_leaf` is true when the node has no outgoing dependency edges (value-only
    cells and literal-only formulas such as `=1+1`).

    `is_array_formula` is true when extraction observed a fastpyxl
    `ArrayFormula` on this cell. `array_formula_ref` is the spill / CSE range
    Excel stored (`ref`, e.g. `E1:E10`), or `None` when that attribute was
    missing. Write-back emits `ArrayFormula` from these fields and refuses to
    scalarize a flagged cell whose `ref` was not observed. fastpyxl does not
    distinguish legacy CSE from dynamic-array spills; both round-trip as
    `t="array"`.

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
    value: Any
    is_leaf: bool
    is_target: bool
    metadata: Mapping[str, Any]
    min_col: str | None
    min_row: int | None
    max_col: str | None
    max_row: int | None
    address: CellKey | None
    formula_ast: AstNode | None = field(repr=False)
    _unparseable_formula: str | None = field(repr=False)
    is_array_formula: bool
    array_formula_ref: str | None

    def __init__(
        self,
        sheet: str | None = None,
        column: str | None = None,
        row: int | None = None,
        formula: str | None = None,
        *,
        value: Any = None,
        is_leaf: bool = True,
        is_target: bool = False,
        metadata: Mapping[str, Any] = EMPTY_METADATA,
        min_col: str | None = None,
        min_row: int | None = None,
        max_col: str | None = None,
        max_row: int | None = None,
        address: CellKey | None = None,
        formula_ast: AstNode | None = None,
        normalized_formula: str | None = None,
        is_array_formula: bool = False,
        array_formula_ref: str | None = None,
    ) -> None:
        self.sheet = sheet
        self.column = column
        self.row = row
        self.formula = formula
        self.value = value
        self.is_leaf = is_leaf
        self.is_target = is_target
        self.metadata = metadata
        self.min_col = min_col
        self.min_row = min_row
        self.max_col = max_col
        self.max_row = max_row
        self.address = address
        self.formula_ast = formula_ast
        self._unparseable_formula = None
        self.is_array_formula = bool(is_array_formula)
        self.array_formula_ref = (
            array_formula_ref if self.is_array_formula and array_formula_ref else None
        )

        if not self.metadata:
            # Never hold a caller's empty dict; share the singleton instead.
            self.metadata = EMPTY_METADATA

        if self.address is not None:
            self._sync_fields_from_address()
        else:
            self._finalize_cell_legacy()
            assert self.sheet is not None and self.column is not None and self.row is not None
            self.address = CellKey(_format_cell_key(self.sheet, self.column, self.row))

        if self.formula_ast is None and normalized_formula:
            ast = parse_formula_text(normalized_formula, anchor=self.address)
            if ast is not None:
                self.formula_ast = ast
            else:
                self._unparseable_formula = normalized_formula

    def _sync_fields_from_address(self) -> None:
        addr = self.address
        assert isinstance(addr, CellKey)
        self.sheet = addr.sheet
        self.column = addr.column
        self.row = addr.row
        self.min_col = addr.column
        self.max_col = addr.column
        self.min_row = addr.row
        self.max_row = addr.row

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
        if self.min_col != col or self.max_col != col or self.min_row != row or self.max_row != row:
            raise ValueError(
                "Graph nodes must be single cells; extent fields must match sheet/column/row"
            )

    def __setattr__(self, name: str, value: Any) -> None:
        if name == "address":
            if value is not None and not isinstance(value, CellKey):
                value = _coerce_cell_address(value)
            current = getattr(self, "address", None)
            if current is not None and value != current:
                raise ValueError(
                    "changing Node.address would retarget relative formula axes; "
                    "use DependencyGraph.move_node"
                )
        object.__setattr__(self, name, value)

    def _relocate(self, new_address: CellKey | str) -> None:
        """Update identity after `DependencyGraph.move_node` rewrites relative axes."""
        object.__setattr__(self, "address", _coerce_cell_address(new_address))
        self._sync_fields_from_address()

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
        return _derived_fields(self.address).column_index

    @property
    def normalized_formula(self) -> str | None:
        """Absolute A1 formula text derived from `formula_ast`, or unparseable fallback."""
        if self.formula_ast is not None:
            return render_formula(
                self.formula_ast,
                anchor=self.address,
                style=FormulaStyle.A1_ABSOLUTE,
            )
        return self._unparseable_formula

    @property
    def has_formula(self) -> bool:
        """True when this cell has a formula AST or unparseable formula text."""
        return self.formula_ast is not None or self._unparseable_formula is not None

    def apply_formula_text(self, text: str | None) -> None:
        """Parse `text` into `formula_ast`, or keep it as unparseable fallback.

        When `address` is set, parse with `parse_preserving_axes` so `$` vs
        bare A1 is kept. Without an anchor, relative intent cannot be
        recovered. Clears both when `text` is missing or blank. Used when a
        rewrite only has formula strings (the AST parser could not handle the
        cell).
        """
        if text is None or not str(text).strip():
            self.formula_ast = None
            self._unparseable_formula = None
            return
        ast = parse_formula_text(text, anchor=self.address)
        self.formula_ast = ast
        self._unparseable_formula = None if ast is not None else text


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
    value: Any
    is_leaf: bool
    is_target: bool
    metadata: Mapping[str, Any]
    min_col: str | None = None
    min_row: int | None = None
    max_col: str | None = None
    max_row: int | None = None
    address: CellKey | None = None
    formula_ast: AstNode | None = field(default=None, repr=False)
    _unparseable_formula: str | None = field(default=None, repr=False)
    is_array_formula: bool = False
    array_formula_ref: str | None = None

    @property
    def key(self) -> NodeKey:
        if self.address is not None:
            return str(self.address)
        assert self.sheet is not None and self.column is not None and self.row is not None
        return _format_cell_key(self.sheet, self.column, self.row)

    @property
    def shape(self) -> NodeShape:
        return NodeShape.cell

    @property
    def column_index(self) -> int:
        if self.column is None:
            raise ValueError("column_index requires a column")
        return int(fastpyxl.utils.cell.column_index_from_string(self.column))

    @property
    def normalized_formula(self) -> str | None:
        """Absolute A1 formula text derived from `formula_ast`, or unparseable fallback."""
        if self.formula_ast is not None:
            return render_formula(
                self.formula_ast,
                anchor=self.address,
                style=FormulaStyle.A1_ABSOLUTE,
            )
        return self._unparseable_formula

    @property
    def has_formula(self) -> bool:
        """True when this cell has a formula AST or unparseable formula text."""
        return self.formula_ast is not None or self._unparseable_formula is not None


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
        value=node.value,
        is_leaf=node.is_leaf,
        is_target=node.is_target,
        metadata=_view_metadata(node.metadata),
        min_col=node.min_col,
        min_row=node.min_row,
        max_col=node.max_col,
        max_row=node.max_row,
        address=node.address,
        formula_ast=node.formula_ast,
        _unparseable_formula=node._unparseable_formula,
        is_array_formula=node.is_array_formula,
        array_formula_ref=node.array_formula_ref,
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
    formula_ast: AstNode | None = None,
    is_array_formula: bool = False,
    array_formula_ref: str | None = None,
) -> Node:
    """Build a single-cell graph node.

    Prefer `formula_ast=` when the tree is already known. `normalized_formula`
    is bootstrap text parsed with axis intent against this cell's address.
    `is_array_formula` / `array_formula_ref` capture CSE or spill provenance
    from extraction; omit them for ordinary scalar formulas.
    """
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
        formula_ast=formula_ast,
        is_array_formula=is_array_formula,
        array_formula_ref=array_formula_ref,
    )


def copy_node(node: Node) -> Node:
    """Return a field-wise copy of `node`.

    Metadata is copied; `formula_ast` is shared because AST nodes are frozen.
    """
    return Node(
        sheet=node.sheet,
        column=node.column,
        row=node.row,
        formula=node.formula,
        normalized_formula=None if node.formula_ast is not None else node._unparseable_formula,
        value=node.value,
        is_leaf=node.is_leaf,
        is_target=node.is_target,
        metadata=copy_metadata(node.metadata),
        min_col=node.min_col,
        min_row=node.min_row,
        max_col=node.max_col,
        max_row=node.max_row,
        address=node.address,
        formula_ast=node.formula_ast,
        is_array_formula=node.is_array_formula,
        array_formula_ref=node.array_formula_ref,
    )
