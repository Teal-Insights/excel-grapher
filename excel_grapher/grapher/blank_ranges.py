"""Sheet-qualified rectangular regions treated as structurally empty when building graphs."""

from __future__ import annotations

import importlib.util
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import TypeAlias, cast

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import normalize_key, parse_address, parse_cell_coords

BlankRangeRect: TypeAlias = tuple[str, int, int, int, int]  # sheet, r1, c1, r2, c2 inclusive


def parse_blank_range_spec(spec: str) -> BlankRangeRect:
    """Parse a sheet-qualified A1 range into normalized inclusive bounds."""
    if not isinstance(spec, str):
        raise TypeError("blank range spec must be a string")
    sheet, cell_part = parse_address(spec)
    if ":" in cell_part:
        start_cell, end_cell = cell_part.split(":", 1)
    else:
        start_cell = end_cell = cell_part

    start_col_str, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell.strip())
    end_col_str, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell.strip())
    start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col_str)
    end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col_str)

    r1, r2 = (start_row, end_row) if start_row <= end_row else (end_row, start_row)
    c1, c2 = (
        (start_col_idx, end_col_idx)
        if start_col_idx <= end_col_idx
        else (end_col_idx, start_col_idx)
    )
    return (sheet, r1, c1, r2, c2)


def normalize_blank_range_specs(specs: Iterable[str] | None) -> tuple[BlankRangeRect, ...]:
    """Normalize a sequence of sheet-qualified range strings."""
    if specs is None:
        return ()
    if isinstance(specs, (str, bytes)):
        raise TypeError("blank_ranges must be a sequence of strings, not a single string")
    return tuple(parse_blank_range_spec(str(s)) for s in specs)


# Identity cache: blank-rect tuples are immutable, so the sheet groups stay valid
# for the life of the tuple. Lists are grouped per call and not cached.
_SHEET_INDEX: dict[
    int, tuple[tuple[BlankRangeRect, ...], dict[str, tuple[BlankRangeRect, ...]]]
] = {}
_SHEET_INDEX_LIMIT = 16


def _rects_by_sheet(
    rects: Sequence[BlankRangeRect],
) -> dict[str, tuple[BlankRangeRect, ...]]:
    """Group `rects` by sheet.

    Repeated lookups of the same tuple reuse one grouping. The returned
    mapping is cached for tuples; callers must not mutate it.
    """
    if isinstance(rects, tuple):
        rect_tuple = cast("tuple[BlankRangeRect, ...]", rects)
        cached = _SHEET_INDEX.get(id(rect_tuple))
        if cached is not None and cached[0] is rect_tuple:
            return cached[1]
    else:
        rect_tuple = None
    grouped: dict[str, list[BlankRangeRect]] = {}
    for rect in rects:
        grouped.setdefault(rect[0], []).append(rect)
    frozen = {sheet: tuple(items) for sheet, items in grouped.items()}
    if rect_tuple is not None:
        if len(_SHEET_INDEX) >= _SHEET_INDEX_LIMIT:
            _SHEET_INDEX.clear()
        _SHEET_INDEX[id(rect_tuple)] = (rect_tuple, frozen)
    return frozen


def cell_in_blank_ranges(sheet: str, row: int, col: int, rects: Sequence[BlankRangeRect]) -> bool:
    """True if (sheet, row, col) lies in any declared blank rectangle."""
    for _sh, r1, c1, r2, c2 in _rects_by_sheet(rects).get(sheet, ()):
        if r1 <= row <= r2 and c1 <= col <= c2:
            return True
    return False


def _merge_spans(spans: list[tuple[int, int]]) -> tuple[tuple[int, int], ...]:
    """Merge inclusive intervals that touch or overlap."""
    if not spans:
        return ()
    spans.sort()
    merged: list[tuple[int, int]] = [spans[0]]
    for lo, hi in spans[1:]:
        prev_lo, prev_hi = merged[-1]
        if lo <= prev_hi + 1:
            merged[-1] = (prev_lo, max(prev_hi, hi))
        else:
            merged.append((lo, hi))
    return tuple(merged)


def column_in_spans(col: int, spans: tuple[tuple[int, int], ...]) -> bool:
    """True when `col` lies in a sorted, merged list of inclusive intervals."""
    for lo, hi in spans:
        if col < lo:
            return False
        if col <= hi:
            return True
    return False


def row_blank_spans(
    sheet: str,
    row1: int,
    col1: int,
    row2: int,
    col2: int,
    rects: Sequence[BlankRangeRect],
) -> dict[int, tuple[tuple[int, int], ...]]:
    """Merged blank column intervals inside an inclusive rectangle, by row.

    The rectangle is tested against blank rectangles by its corners. Rows
    outside the intersection are omitted. Callers clip queries to this
    rectangle, so a blank rect that extends past it does not allocate the
    exterior.
    """
    if row1 > row2:
        row1, row2 = row2, row1
    if col1 > col2:
        col1, col2 = col2, col1
    if not rects:
        return {}
    pending: dict[int, list[tuple[int, int]]] = {}
    for _sheet, r1, c1, r2, c2 in overlapping_blank_rects(sheet, row1, col1, row2, col2, rects):
        lo_c = c1 if c1 > col1 else col1
        hi_c = c2 if c2 < col2 else col2
        lo_r = r1 if r1 > row1 else row1
        hi_r = r2 if r2 < row2 else row2
        for row in range(lo_r, hi_r + 1):
            pending.setdefault(row, []).append((lo_c, hi_c))
    return {row: _merge_spans(spans) for row, spans in pending.items()}


def address_in_blank_ranges(address: str, rects: Sequence[BlankRangeRect]) -> bool:
    """True if the sheet-qualified cell address falls within any blank range."""
    if not rects:
        return False
    norm = normalize_key(address)
    sheet, cell = parse_address(norm)
    col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
    col = fastpyxl.utils.cell.column_index_from_string(col_str)
    return cell_in_blank_ranges(sheet, int(row), col, rects)


def _rects_overlap(left: BlankRangeRect, right: BlankRangeRect) -> bool:
    """True when two inclusive rectangles share a cell."""
    if left[0] != right[0]:
        return False
    return (
        left[1] <= right[3] and right[1] <= left[3] and left[2] <= right[4] and right[2] <= left[4]
    )


def range_rect(start: str, end: str) -> BlankRangeRect | None:
    """Inclusive same-sheet rectangle for `start:end`, or `None` if sheets differ."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        return None
    r1, r2 = (row1, row2) if row1 <= row2 else (row2, row1)
    c1, c2 = (col1, col2) if col1 <= col2 else (col2, col1)
    return (sheet1, r1, c1, r2, c2)


def range_overlaps_blank_ranges(start: str, end: str, rects: Sequence[BlankRangeRect]) -> bool:
    """True if the same-sheet rectangle `start:end` overlaps any blank rect."""
    if not rects:
        return False
    probe = range_rect(start, end)
    if probe is None:
        return False
    return any(_rects_overlap(probe, rect) for rect in _rects_by_sheet(rects).get(probe[0], ()))


def overlapping_blank_rects(
    sheet: str,
    row1: int,
    col1: int,
    row2: int,
    col2: int,
    rects: Sequence[BlankRangeRect],
) -> tuple[BlankRangeRect, ...]:
    """Blank rectangles that geometrically overlap the given inclusive bounds."""
    if not rects:
        return ()
    probe = (sheet, row1, col1, row2, col2)
    return tuple(
        rect for rect in _rects_by_sheet(rects).get(sheet, ()) if _rects_overlap(probe, rect)
    )


def blank_rects_for_addresses(
    addresses: Sequence[str],
    rects: Sequence[BlankRangeRect],
) -> tuple[BlankRangeRect, ...]:
    """Blank rectangles that overlap the bounding box of `addresses`."""
    if not addresses or not rects:
        return ()
    bounds: dict[str, tuple[int, int, int, int]] = {}
    for address in addresses:
        sheet, row, col = parse_cell_coords(address)
        previous = bounds.get(sheet)
        if previous is None:
            bounds[sheet] = (row, col, row, col)
            continue
        r1, c1, r2, c2 = previous
        bounds[sheet] = (min(r1, row), min(c1, col), max(r2, row), max(c2, col))
    found: list[BlankRangeRect] = []
    for sheet, (row1, col1, row2, col2) in bounds.items():
        found.extend(overlapping_blank_rects(sheet, row1, col1, row2, col2, rects))
    return tuple(found)


class BlankRangesLoadError(ValueError):
    """Raised when a `BLANK_RANGES` module cannot be loaded."""


def load_blank_ranges_module(path: Path | str) -> tuple[str, ...]:
    """Import a module exposing `BLANK_RANGES: Sequence[str]`.

    Args:
        path: Filesystem path to a Python module.

    Returns:
        The module's `BLANK_RANGES` sequence as strings.

    Raises:
        BlankRangesLoadError: When the file is missing, cannot be imported, or
            does not expose `BLANK_RANGES` as a sequence of strings.
    """
    resolved = Path(path).resolve()
    if not resolved.is_file():
        raise BlankRangesLoadError(f"Blank-ranges module not found: {resolved}")
    spec = importlib.util.spec_from_file_location(
        f"excel_grapher_blank_ranges_{resolved.stem}",
        resolved,
    )
    if spec is None or spec.loader is None:
        raise BlankRangesLoadError(f"Cannot load blank-ranges module: {resolved}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        raise BlankRangesLoadError(
            f"Failed to import blank-ranges module {resolved}: {exc}"
        ) from exc
    table = getattr(module, "BLANK_RANGES", None)
    if table is None:
        raise BlankRangesLoadError(
            f"Blank-ranges module {resolved} must define BLANK_RANGES: Sequence[str]"
        )
    if isinstance(table, (str, bytes)) or not isinstance(table, Sequence):
        raise BlankRangesLoadError(
            f"BLANK_RANGES in {resolved} must be a sequence of strings, not {type(table).__name__}"
        )
    return tuple(str(item) for item in table)
