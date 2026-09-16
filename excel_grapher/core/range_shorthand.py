"""Workbook-aware resolution for Excel whole-column and whole-row range shorthand.

Excel allows `C:C` (entire column) and `5:5` (entire row) in formulas, and the
same shorthand as defined-name referents (`Sheet!$5:$134`, `Sheet!$A:$C`). We
parse those forms syntactically and resolve them to bounded `ExcelRange` values
using each sheet's **used range** from the workbook (`max_row`, `max_col`), not
Excel's full grid (`XFD` / `1048576`). Rectangular ranges that exceed
`max_range_cells` raise `ValueError` (fail closed); whole-column/row shorthands
always expand to every cell in the used-range extent so `MATCH`/`INDEX` on
interior rows remain correct.
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


def resolve_whole_column_span(
    sheet: str, start_col: str, end_col: str, bounds: SheetBounds
) -> ExcelRange:
    """Map a whole-column span (`A:C`) to a used-range-bounded `ExcelRange`.

    Rows run from 1 through the sheet's used `max_row` (Excel's last row when
    *sheet* is absent from *bounds*). Named columns are kept; only the implicit
    row axis is filled from used-range bounds.
    """
    c1 = column_index_from_string(start_col.upper())
    c2 = column_index_from_string(end_col.upper())
    start_c, end_c = min(c1, c2), max(c1, c2)
    max_r, _ = sheet_used_extent(bounds, sheet)
    return ExcelRange(
        sheet=sheet,
        start_row=1,
        start_col=start_c,
        end_row=max_r,
        end_col=end_c,
    )


def resolve_whole_row_span(
    sheet: str, start_row: int, end_row: int, bounds: SheetBounds
) -> ExcelRange:
    """Map a whole-row span (`5:10`) to a used-range-bounded `ExcelRange`.

    Columns run from 1 through the sheet's used `max_col` (Excel's last column
    when *sheet* is absent from *bounds*). Named rows are kept; only the
    implicit column axis is filled from used-range bounds.
    """
    r1, r2 = min(start_row, end_row), max(start_row, end_row)
    _, max_c = sheet_used_extent(bounds, sheet)
    return ExcelRange(
        sheet=sheet,
        start_row=r1,
        start_col=1,
        end_row=r2,
        end_col=max_c,
    )


def resolve_whole_column(sheet: str, column: str, bounds: SheetBounds) -> ExcelRange:
    """Map a whole-column shorthand to a bounded single-column `ExcelRange`."""
    return resolve_whole_column_span(sheet, column, column, bounds)


def resolve_whole_row(sheet: str, row: int, bounds: SheetBounds) -> ExcelRange:
    """Map a whole-row shorthand to a bounded single-row `ExcelRange`."""
    return resolve_whole_row_span(sheet, row, row, bounds)


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
