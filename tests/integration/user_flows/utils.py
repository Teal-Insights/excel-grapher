from __future__ import annotations

from collections.abc import Callable
from itertools import count
from pathlib import Path

import xlsxwriter
from xlsxwriter.worksheet import Worksheet

CellValue = int | float | str | None
WorkbookBuilder = Callable[[Worksheet, xlsxwriter.Workbook], None]
WorkbookFactory = Callable[[WorkbookBuilder], Path]


def build_workbook_factory(tmp_path: Path, *, prefix: str) -> WorkbookFactory:
    seq = count()

    def _build(builder: WorkbookBuilder) -> Path:
        path = tmp_path / f"{prefix}_{next(seq)}.xlsx"
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet("Sheet1")
        builder(worksheet, workbook)
        workbook.close()
        return path

    return _build


def write_single_row(
    worksheet: Worksheet,
    row_cells: tuple[CellValue, ...],
    *,
    row_idx: int = 0,
) -> None:
    for col_idx, cell in enumerate(row_cells):
        if isinstance(cell, str) and cell.startswith("="):
            worksheet.write_formula(row_idx, col_idx, cell, None, 0)
        elif isinstance(cell, str):
            worksheet.write_string(row_idx, col_idx, cell)
        elif isinstance(cell, (int, float)):
            worksheet.write_number(row_idx, col_idx, cell)


def write_series_bindings_workbook_blocks(
    worksheet: Worksheet,
    _workbook: xlsxwriter.Workbook,
) -> None:
    worksheet.write("D1", "Year")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        worksheet.write_number(0, col, year)
    worksheet.write("A2", "Borvelia")
    worksheet.write("A3", "Real GDP growth (% per annum)")
    worksheet.write("A4", "Real interest rate (% per annum)")
    worksheet.write("A5", "Primary balance (% of GDP)")
    for col, value in enumerate([2.1, 2.2, 2.3, 2.4, 2.5], start=5):
        worksheet.write_number(2, col, value)
    for col, value in enumerate([1.0, 1.1, 1.2, 1.3, 1.4], start=5):
        worksheet.write_number(3, col, value)
    for col, value in enumerate([-1.0, -0.5, 0.0, 0.5, 1.0], start=5):
        worksheet.write_number(4, col, value)


def write_series_bindings_workbook(path: Path) -> None:
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Sheet1")
    write_series_bindings_workbook_blocks(worksheet, workbook)
    workbook.close()
