"""Shared helpers for explicit matrix layout binding tests."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

FIXTURES = Path(__file__).resolve().parent
MATRIX_EXPLICIT_BINDINGS = FIXTURES / "matrix_explicit_1_4_0.yaml"


def write_matrix_explicit_workbook(path: Path) -> None:
    """Build a small rectangular matrix workbook for macro_matrix bindings."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Indicator")
    periods = [2024, 2025, 2026]
    for col_offset, period in enumerate(periods):
        ws.write_number(1, 1 + col_offset, period)
    rows = [
        ("GDP growth", [1.2, 1.4, 1.5]),
        ("Inflation", [3.1, 2.9, 2.7]),
        ("Debt", [55.0, 54.2, 53.8]),
    ]
    for row_offset, (indicator, values) in enumerate(rows):
        excel_row = 2 + row_offset
        ws.write(excel_row, 0, indicator)
        for col_offset, value in enumerate(values):
            ws.write_number(excel_row, 1 + col_offset, value)
    wb.close()
