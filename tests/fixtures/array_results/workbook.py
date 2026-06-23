"""Builders and paths for top-level array-result workbook fixtures (#284)."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

ARRAY_RESULTS_FIXTURE_DIR = Path(__file__).resolve().parent

COLUMN_COMPARE_XLSX = "column_compare.xlsx"
ROW_COMPARE_XLSX = "row_compare.xlsx"
NUMERIC_COMPARE_XLSX = "numeric_compare.xlsx"
BLOCKED_SPILL_XLSX = "blocked_spill.xlsx"
GRID_2D_COMPARE_XLSX = "grid_2d_compare.xlsx"


def column_compare_path() -> Path:
    return ARRAY_RESULTS_FIXTURE_DIR / COLUMN_COMPARE_XLSX


def row_compare_path() -> Path:
    return ARRAY_RESULTS_FIXTURE_DIR / ROW_COMPARE_XLSX


def numeric_compare_path() -> Path:
    return ARRAY_RESULTS_FIXTURE_DIR / NUMERIC_COMPARE_XLSX


def blocked_spill_path() -> Path:
    return ARRAY_RESULTS_FIXTURE_DIR / BLOCKED_SPILL_XLSX


def grid_2d_compare_path() -> Path:
    return ARRAY_RESULTS_FIXTURE_DIR / GRID_2D_COMPARE_XLSX


def build_column_compare_workbook(path: Path) -> Path:
    """``Data!D10 = C5:C7="Software"`` — column vector of booleans."""
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Data")
    for row_offset, category in enumerate(["Software", "Hardware", "Software"]):
        worksheet.write_string(4 + row_offset, 2, category)
    worksheet.write_formula(9, 3, '=C5:C7="Software"')
    workbook.close()
    return path


def build_row_compare_workbook(path: Path) -> Path:
    """``Data!F5 = C5:E5="A"`` — row vector of booleans."""
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Data")
    for col_offset, label in enumerate(["A", "B", "A"]):
        worksheet.write_string(4, 2 + col_offset, label)
    worksheet.write_formula(4, 5, '=C5:E5="A"')
    workbook.close()
    return path


def build_numeric_compare_workbook(path: Path) -> Path:
    """``Data!D10 = C5:C7>E5:E7`` — column vector of numeric comparisons."""
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Data")
    left = [10.0, 5.0, 8.0]
    right = [3.0, 7.0, 8.0]
    for row_offset, (left_value, right_value) in enumerate(zip(left, right, strict=True)):
        worksheet.write_number(4 + row_offset, 2, left_value)
        worksheet.write_number(4 + row_offset, 4, right_value)
    worksheet.write_formula(9, 3, "=C5:C7>E5:E7")
    workbook.close()
    return path


def build_blocked_spill_workbook(path: Path) -> Path:
    """``Data!D10 = C5:C7="Software"`` with ``Data!D11`` occupied → ``#SPILL!``."""
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Data")
    for row_offset, category in enumerate(["Software", "Hardware", "Software"]):
        worksheet.write_string(4 + row_offset, 2, category)
    worksheet.write_number(10, 3, 1)
    worksheet.write_formula(9, 3, '=C5:C7="Software"', None, "#SPILL!")
    workbook.close()
    return path


def build_grid_2d_compare_workbook(path: Path) -> Path:
    """``Data!D10 = C5:D6>5`` — 2×2 boolean array at top level."""
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Data")
    values = ((10.0, 3.0), (8.0, 12.0))
    for row_offset, row_values in enumerate(values):
        for col_offset, value in enumerate(row_values):
            worksheet.write_number(4 + row_offset, 2 + col_offset, value)
    worksheet.write_formula(9, 3, "=C5:D6>5")
    workbook.close()
    return path


def ensure_committed_fixtures() -> None:
    """Write static xlsx fixtures when missing (e.g. fresh checkout)."""
    builders = (
        (column_compare_path(), build_column_compare_workbook),
        (row_compare_path(), build_row_compare_workbook),
        (numeric_compare_path(), build_numeric_compare_workbook),
        (blocked_spill_path(), build_blocked_spill_workbook),
        (grid_2d_compare_path(), build_grid_2d_compare_workbook),
    )
    for path, builder in builders:
        if not path.is_file():
            builder(path)
