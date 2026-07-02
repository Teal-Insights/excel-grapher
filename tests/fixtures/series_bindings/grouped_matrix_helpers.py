"""Shared helpers for grouped-row matrix layout binding tests.

Models the "Discrete Risks" band pattern: a group label (SCENARIO) appears
once per band in column A, sub-rows carry a second row dimension (SHOCK_TYPE)
in the same column, and a decorative blank row separates bands. A SUM formula
references the whole block so blank rows become graph leaves, exercising
`exclude_rows`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import xlsxwriter

FIXTURES = Path(__file__).resolve().parent
MATRIX_GROUPED_ROWS_BINDINGS = FIXTURES / "matrix_grouped_rows_1_5_0.yaml"

GROUPED_MATRIX_PERIODS: tuple[int, ...] = (2024, 2025)

# (row, scenario, shock_type, values) for data rows only.
GROUPED_MATRIX_DATA_ROWS: tuple[tuple[int, str, str, list[float]], ...] = (
    (3, "Paris", "Revenue", [1.1, 1.2]),
    (4, "Paris", "Primary expenditure", [2.1, 2.2]),
    (7, "Moderate", "Revenue", [3.1, 3.2]),
    (8, "Moderate", "Primary expenditure", [4.1, 4.2]),
)

BAND_HEADER_ROWS: tuple[tuple[int, str], ...] = ((2, "Paris"), (6, "Moderate"))
SEPARATOR_ROW = 5


def write_grouped_matrix_workbook(path: Path) -> None:
    """Build a banded workbook: header rows in column A, blank separator row.

    Layout (data range `Inputs!C2:D8`):

    ```text
            A                     C       D
    1                             2024    2025
    2       Paris
    3       Revenue               1.1     1.2
    4       Primary expenditure   2.1     2.2
    5
    6       Moderate
    7       Revenue               3.1     3.2
    8       Primary expenditure   4.1     4.2
    ```

    `F1` sums the whole block, so even blank cells inside the data range are
    referenced and appear as graph leaves.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    for col_offset, period in enumerate(GROUPED_MATRIX_PERIODS):
        ws.write_number(0, 2 + col_offset, period)
    for row, label in BAND_HEADER_ROWS:
        ws.write(row - 1, 0, label)
    for row, _scenario, shock_type, values in GROUPED_MATRIX_DATA_ROWS:
        ws.write(row - 1, 0, shock_type)
        for col_offset, value in enumerate(values):
            ws.write_number(row - 1, 2 + col_offset, value)
    total = sum(sum(values) for _r, _s, _k, values in GROUPED_MATRIX_DATA_ROWS)
    ws.write_formula("F1", "=SUM(C2:D8)", None, total)
    wb.close()


def grouped_matrix_structure() -> dict[str, Any]:
    """Return the grouped-row matrix structure block (banded column A)."""
    return {
        "measure": {
            "concept": "OBS_VALUE",
            "dtype": "float",
            "bind": {"kind": "data_cell", "read": "float"},
        },
        "dimensions": [
            {
                "concept": "SCENARIO",
                "role": "key",
                "scope": "cell",
                "bind": {
                    "kind": "row_label",
                    "label_column": "A",
                    "include": [2, 6],
                    "fill": True,
                    "read": "string",
                },
            },
            {
                "concept": "SHOCK_TYPE",
                "role": "key",
                "scope": "cell",
                "bind": {
                    "kind": "row_label",
                    "label_column": "A",
                    "skip": [2, 6],
                    "read": "string",
                },
            },
            {
                "concept": "TIME_PERIOD",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
            },
        ],
    }


def grouped_matrix_series(**overrides: Any) -> dict[str, Any]:
    """Return an inline grouped-row matrix series entry for tests."""
    series: dict[str, Any] = {
        "id": "discrete_risks",
        "sheet": "Inputs",
        "data_range": "Inputs!C2:D8",
        "layout": "matrix",
        "exclude_rows": [2, "5:6"],
        "input": {"setter": {"name": "set_discrete_risks"}},
        "structure": grouped_matrix_structure(),
        "key": ["SCENARIO", "SHOCK_TYPE", "TIME_PERIOD"],
    }
    series.update(overrides)
    return series


def grouped_matrix_bindings_document(**series_overrides: Any) -> dict[str, Any]:
    """Return a binding manifest document for grouped-row matrix tests."""
    return {
        "schema_version": "1.5.0",
        "workbook": "grouped_inputs.xlsx",
        "series": [grouped_matrix_series(**series_overrides)],
    }


def expected_grouped_matrix_keys() -> set[tuple[str, str, int]]:
    """Expected (SCENARIO, SHOCK_TYPE, TIME_PERIOD) composite keys."""
    return {
        (scenario, shock_type, period)
        for _row, scenario, shock_type, _values in GROUPED_MATRIX_DATA_ROWS
        for period in GROUPED_MATRIX_PERIODS
    }
