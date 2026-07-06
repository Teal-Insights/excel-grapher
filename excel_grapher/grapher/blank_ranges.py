"""Sheet-qualified rectangular regions treated as structurally empty when building graphs."""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from typing import TypeAlias

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import normalize_key, parse_address

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


def cell_in_blank_ranges(sheet: str, row: int, col: int, rects: Sequence[BlankRangeRect]) -> bool:
    """True if (sheet, row, col) lies in any declared blank rectangle."""
    for sh, r1, c1, r2, c2 in rects:
        if sh != sheet:
            continue
        if r1 <= row <= r2 and c1 <= col <= c2:
            return True
    return False


def address_in_blank_ranges(address: str, rects: Sequence[BlankRangeRect]) -> bool:
    """True if the sheet-qualified cell address falls within any blank range."""
    if not rects:
        return False
    norm = normalize_key(address)
    sheet, cell = parse_address(norm)
    col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
    col = fastpyxl.utils.cell.column_index_from_string(col_str)
    return cell_in_blank_ranges(sheet, int(row), col, rects)
