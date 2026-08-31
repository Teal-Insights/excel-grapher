"""Sparse coordinate store for exported leaf values.

Generated `DEFAULT_INPUTS` uses nested `dict[sheet, dict[(row, col), value]]`
maps. Generated `CONSTANTS` uses the same shape wrapped in `MappingProxyType`
at both levels so constant leaves are not an overridable input table. Public
APIs (`make_context`, `set_inputs`) still accept NodeKey strings and fail
closed when a key cannot round-trip to `(sheet, row, col)`.
"""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any, TypeAlias, cast

from excel_grapher.core import CellValue
from excel_grapher.core.address_keys import parse_cell_coords

__all__ = [
    "MISSING",
    "LeafInputs",
    "LeafStore",
    "as_leaf_store",
    "leaf",
    "lookup_leaf",
    "overlay_leaf_inputs",
    "prepare_context_inputs",
]

LeafStore: TypeAlias = dict[str, dict[tuple[int, int], CellValue]]

MISSING: object = object()


class LeafInputs:
    """NodeKey-keyed view over a nested `LeafStore`.

    Get/set parse A1 at the boundary. Rectangle scans should call `leaf`
    with integer coordinates instead of iterating this view.
    """

    __slots__ = ("_store",)

    def __init__(self, store: LeafStore) -> None:
        self._store = store

    def __getitem__(self, address: str) -> CellValue:
        sheet, row, col = _require_cell_coords(address)
        try:
            return self._store[sheet][(row, col)]
        except KeyError:
            raise KeyError(address) from None

    def __setitem__(self, address: str, value: CellValue) -> None:
        sheet, row, col = _require_cell_coords(address)
        self._store.setdefault(sheet, {})[(row, col)] = value

    def __contains__(self, address: object) -> bool:
        if not isinstance(address, str):
            return False
        try:
            sheet, row, col = parse_cell_coords(address)
        except ValueError:
            return False
        sheet_map = self._store.get(sheet)
        if sheet_map is None:
            return False
        return (row, col) in sheet_map

    def get(self, address: str, default: CellValue | None = None) -> CellValue | None:
        try:
            return self[address]
        except KeyError:
            return default

    def __eq__(self, other: object) -> bool:
        if isinstance(other, LeafInputs):
            return self._store == other._store
        return NotImplemented


def leaf(store: LeafStore, sheet: str, row: int, col: int) -> object:
    """Return the stored leaf at `(sheet, row, col)`, or `MISSING` if absent."""
    sheet_map = store.get(sheet)
    if sheet_map is None:
        return MISSING
    return sheet_map.get((row, col), MISSING)


def lookup_leaf(ctx: object, address: str) -> Any:
    """Look up a leaf by NodeKey, using `ctx.leaves` when present.

    Returns `MISSING` when the address is not a stored leaf (including when it
    is not a parseable single cell). Formula cells must still go through the
    resolver.
    """
    store = getattr(ctx, "leaves", None)
    if not store:
        return MISSING
    try:
        sheet, row, col = parse_cell_coords(address)
    except ValueError:
        return MISSING
    return leaf(cast(LeafStore, store), sheet, row, col)


def as_leaf_store(values: Any) -> LeafStore:
    """Copy `values` into a nested leaf store.

    Accepts a nested coordinate store or a NodeKey-keyed dict. Empty mappings
    become `{}`.
    """
    if isinstance(values, LeafInputs):
        return {sheet: dict(cells) for sheet, cells in values._store.items()}
    if not isinstance(values, Mapping):
        raise TypeError(f"Expected a mapping of leaves; got {type(values)!r}")
    if not values:
        return {}
    if _looks_like_leaf_store(values):
        out: LeafStore = {}
        for sheet, cells in values.items():
            if not isinstance(sheet, str):
                raise TypeError(f"Leaf store sheets must be strings; got {sheet!r}")
            if not isinstance(cells, Mapping):
                raise TypeError(f"Leaf store sheet map must be a mapping; got {cells!r}")
            sheet_map: dict[tuple[int, int], CellValue] = {}
            for coord, value in cells.items():
                row, col = _require_coord(coord)
                sheet_map[(row, col)] = cast(CellValue, value)
            out[sheet] = sheet_map
        return out
    out = {}
    overlay_leaf_inputs(out, values)
    return out


def overlay_leaf_inputs(store: LeafStore, overlay: Any) -> None:
    """Merge `overlay` into `store`.

    Nested coordinate stores merge sheet-by-sheet. NodeKey dicts parse A1 at
    the boundary.

    Raises:
        ValueError: If a NodeKey cannot round-trip to `(sheet, row, col)`.
        TypeError: If `overlay` is neither a leaf store nor a NodeKey mapping.
    """
    if not overlay:
        return
    if not isinstance(overlay, Mapping):
        raise TypeError(f"Expected a mapping of leaves; got {type(overlay)!r}")
    if _looks_like_leaf_store(overlay):
        for sheet, cells in overlay.items():
            if not isinstance(sheet, str):
                raise TypeError(f"Leaf store sheets must be strings; got {sheet!r}")
            if not isinstance(cells, Mapping):
                raise TypeError(f"Leaf store sheet map must be a mapping; got {cells!r}")
            sheet_map = store.setdefault(sheet, {})
            for coord, value in cells.items():
                row, col = _require_coord(coord)
                sheet_map[(row, col)] = cast(CellValue, value)
        return
    for key, value in overlay.items():
        if not isinstance(key, str):
            raise TypeError(f"Input keys must be NodeKey strings; got {key!r}")
        sheet, row, col = _require_cell_coords(key)
        store.setdefault(sheet, {})[(row, col)] = cast(CellValue, value)


def prepare_context_inputs(
    default_inputs: Any,
    constants: Any = None,
    overlay: Any = None,
) -> LeafStore:
    """Copy defaults, then merge constants and a NodeKey overlay."""
    merged = as_leaf_store(default_inputs)
    if constants:
        overlay_leaf_inputs(merged, constants)
    if overlay is not None:
        overlay_leaf_inputs(merged, overlay)
    return merged


def _require_cell_coords(address: str) -> tuple[str, int, int]:
    try:
        return parse_cell_coords(address)
    except ValueError as exc:
        raise ValueError(f"Cannot round-trip input key to (sheet, row, col): {address!r}") from exc


def _require_coord(coord: object) -> tuple[int, int]:
    if not (isinstance(coord, tuple) and len(coord) == 2):
        raise TypeError(f"Leaf store keys must be (row, col) tuples; got {coord!r}")
    row, col = coord
    if not isinstance(row, int) or not isinstance(col, int):
        raise TypeError(f"Leaf store keys must be (row, col) ints; got {coord!r}")
    if row < 1 or col < 1:
        raise ValueError(f"Leaf store coordinates must be 1-based; got {coord!r}")
    return row, col


def _looks_like_leaf_store(values: Mapping[object, object]) -> bool:
    """True when values look like `sheet -> {(row, col): value}`."""
    for sample in values.values():
        if not isinstance(sample, Mapping):
            return False
        for coord in sample:
            return (
                isinstance(coord, tuple)
                and len(coord) == 2
                and isinstance(coord[0], int)
                and isinstance(coord[1], int)
            )
        return True
    return False
