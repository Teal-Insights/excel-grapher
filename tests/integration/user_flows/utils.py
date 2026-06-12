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


def write_calendar_flags_workbook(path: Path) -> None:
    """Build a small workbook with bool row flags and datetime column headers."""
    from datetime import datetime

    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Inputs")
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
    periods = [
        datetime(2024, 1, 1),
        datetime(2024, 2, 1),
        datetime(2024, 3, 1),
    ]
    for col_index, period in enumerate(periods, start=1):
        worksheet.write_datetime(0, col_index, period, date_format)
    rows = [
        (True, (10.0, 20.0, 30.0)),
        (False, (5.0, 15.0, 25.0)),
    ]
    for row_index, (is_active, values) in enumerate(rows, start=1):
        worksheet.write_boolean(row_index, 0, is_active)
        for col_index, value in enumerate(values, start=1):
            worksheet.write_number(row_index, col_index, value)
    workbook.close()


def write_ffv2_workbook(path: Path) -> None:
    """Build ffv2.xlsx — game log with datetime headers in row 1 (B:Q)."""
    from datetime import datetime, timedelta

    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Sheet1")
    date_format = workbook.add_format({"num_format": "m/d/yyyy"})

    start_date = datetime(2026, 9, 7)
    column_count = 16
    for offset in range(column_count):
        col_index = offset + 1
        game_date = start_date + timedelta(days=offset)
        worksheet.write_datetime(0, col_index, game_date, date_format)
        worksheet.write_string(1, col_index, f"W{10 + offset}-{9 + offset}")

    row_values: dict[int, list[float]] = {
        2: [10, 8, 11, 13, 10, 12, 9, 13, 11, 11, 7, 6, 5, 8, 10, 8],
        3: [11, 9, 15, 14, 12, 13, 10, 16, 14, 14, 9, 8, 7, 10, 12, 10],
        4: [130, 91, 112, 143, 110, 132, 95, 148, 120, 118, 84, 72, 65, 96, 125, 98],
        5: [13, 11.4, 10.2, 11.9, 11, 11, 10.6, 11.4, 10.9, 10.7, 12, 12, 13, 12, 12.5, 12.3],
        6: [0, 0, 0, 1, 0, 1, 0, 1, 0, 0, 1, 0, 0, 1, 0, 1],
        7: [25, 24, 20, 28, 22, 26, 21, 30, 23, 24, 19, 18, 17, 22, 27, 21],
        17: [
            23,
            17.1,
            22.2,
            36,
            19.5,
            28.4,
            16.8,
            31.2,
            20.1,
            21.0,
            18.3,
            15.6,
            14.2,
            24.5,
            26.8,
            22.0,
        ],
    }
    for row_index, values in row_values.items():
        for col_index, value in enumerate(values, start=1):
            worksheet.write_number(row_index, col_index, value)
    workbook.close()
