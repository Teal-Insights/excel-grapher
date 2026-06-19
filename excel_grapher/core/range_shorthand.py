"""Workbook-aware resolution for Excel whole-column and whole-row range shorthand.

Excel allows ``C:C`` (entire column) and ``5:5`` (entire row) in formulas. We parse
those forms syntactically and resolve them to bounded ``ExcelRange`` values using
each sheet's **used range** from the workbook (``max_row``, ``max_col``), not
Excel's full grid (1:1 048 576). Rectangular ranges that exceed ``max_range_cells``
still collapse to corner endpoints (issue #56); whole-column/row shorthands always
expand to every cell in the used-range extent so ``MATCH``/``INDEX`` on interior
rows remain correct.
"""

from __future__ import annotations

from collections.abc import Iterator

import fastpyxl.utils.cell
from fastpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_grapher.core.address_keys import format_cell_key
from excel_grapher.core.types import ExcelRange

EXCEL_MAX_ROW = 1_048_576
EXCEL_MAX_COL = 16_384

SheetBounds = dict[str, tuple[int, int]]


def sheet_used_extent(bounds: SheetBounds, sheet: str) -> tuple[int, int]:
    """Return ``(max_row, max_col)`` for *sheet*, defaulting to Excel limits."""
    return bounds.get(sheet, (EXCEL_MAX_ROW, EXCEL_MAX_COL))


def resolve_whole_column(sheet: str, column: str, bounds: SheetBounds) -> ExcelRange:
    """Map a whole-column shorthand to a bounded single-column ``ExcelRange``."""
    col = column.upper()
    col_idx = column_index_from_string(col)
    max_r, _ = sheet_used_extent(bounds, sheet)
    return ExcelRange(
        sheet=sheet,
        start_row=1,
        start_col=col_idx,
        end_row=max_r,
        end_col=col_idx,
    )


def resolve_whole_row(sheet: str, row: int, bounds: SheetBounds) -> ExcelRange:
    """Map a whole-row shorthand to a bounded single-row ``ExcelRange``."""
    _, max_c = sheet_used_extent(bounds, sheet)
    return ExcelRange(
        sheet=sheet,
        start_row=row,
        start_col=1,
        end_row=row,
        end_col=max_c,
    )


def iter_whole_column_cells(
    sheet: str, column: str, bounds: SheetBounds
) -> Iterator[tuple[str, str]]:
    """Yield ``(sheet, a1)`` for each cell in a workbook-bounded whole column."""
    rng = resolve_whole_column(sheet, column, bounds)
    col_letter = column.upper()
    for row in range(rng.start_row, rng.end_row + 1):
        yield sheet, f"{col_letter}{row}"


def iter_whole_row_cells(sheet: str, row: int, bounds: SheetBounds) -> Iterator[tuple[str, str]]:
    """Yield ``(sheet, a1)`` for each cell in a workbook-bounded whole row."""
    rng = resolve_whole_row(sheet, row, bounds)
    for col_idx in range(rng.start_col, rng.end_col + 1):
        yield sheet, f"{fastpyxl.utils.cell.get_column_letter(col_idx)}{row}"


def expand_whole_column_deps(sheet: str, column: str, bounds: SheetBounds) -> list[tuple[str, str]]:
    """Expand a whole-column shorthand to all ``(sheet, a1)`` deps in the used range."""
    return list(iter_whole_column_cells(sheet, column, bounds))


def expand_whole_row_deps(sheet: str, row: int, bounds: SheetBounds) -> list[tuple[str, str]]:
    """Expand a whole-row shorthand to all ``(sheet, a1)`` deps in the used range."""
    return list(iter_whole_row_cells(sheet, row, bounds))


def whole_column_to_bounded_a1(sheet: str, column: str, bounds: SheetBounds) -> tuple[str, str]:
    """Return ``(start_ref, end_ref)`` sheet-qualified endpoints for a whole column."""
    col = column.upper()
    max_r, _ = sheet_used_extent(bounds, sheet)
    return (
        format_cell_key(sheet, col, 1),
        format_cell_key(sheet, col, max_r),
    )


def whole_row_to_bounded_a1(sheet: str, row: int, bounds: SheetBounds) -> tuple[str, str]:
    """Return ``(start_ref, end_ref)`` sheet-qualified endpoints for a whole row."""
    _, max_c = sheet_used_extent(bounds, sheet)
    return (
        format_cell_key(sheet, get_column_letter(1), row),
        format_cell_key(sheet, get_column_letter(max_c), row),
    )
