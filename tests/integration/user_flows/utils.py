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
