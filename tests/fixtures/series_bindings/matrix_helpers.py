"""Shared helpers for explicit matrix layout binding tests."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import xlsxwriter

FIXTURES = Path(__file__).resolve().parent
MATRIX_EXPLICIT_BINDINGS = FIXTURES / "matrix_explicit_1_4_0.yaml"
MATRIX_EXPLICIT_COMPUTE_BINDINGS = FIXTURES / "matrix_explicit_compute_1_4_0.yaml"

MACRO_MATRIX_ROWS: tuple[tuple[str, list[float]], ...] = (
    ("GDP growth", [1.2, 1.4, 1.5]),
    ("Inflation", [3.1, 2.9, 2.7]),
    ("Debt", [55.0, 54.2, 53.8]),
)
MACRO_MATRIX_PERIODS: tuple[int, ...] = (2024, 2025, 2026)


def write_matrix_explicit_workbook(path: Path, *, use_formulas: bool = False) -> None:
    """Build a small rectangular matrix workbook for macro_matrix bindings."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Indicator")
    for col_offset, period in enumerate(MACRO_MATRIX_PERIODS):
        ws.write_number(1, 1 + col_offset, period)
    for row_offset, (indicator, values) in enumerate(MACRO_MATRIX_ROWS):
        excel_row = 2 + row_offset
        ws.write(excel_row, 0, indicator)
        for col_offset, value in enumerate(values):
            if use_formulas:
                ws.write_formula(
                    excel_row,
                    1 + col_offset,
                    f"={value}*2",
                    None,
                    float(value) * 2.0,
                )
            else:
                ws.write_number(excel_row, 1 + col_offset, value)
    wb.close()


def macro_matrix_structure() -> dict[str, Any]:
    """Return the shared macro_matrix structure block."""
    return {
        "measure": {
            "concept": "OBS_VALUE",
            "dtype": "float",
            "bind": {"kind": "data_cell", "read": "float"},
        },
        "dimensions": [
            {
                "concept": "INDICATOR",
                "role": "key",
                "scope": "cell",
                "bind": {
                    "kind": "row_label",
                    "label_column": "A",
                    "read": "string",
                    "normalize": "strip",
                },
            },
            {
                "concept": "TIME_PERIOD",
                "role": "key",
                "scope": "cell",
                "bind": {
                    "kind": "column_header",
                    "header_row": 2,
                    "read": "int",
                },
            },
        ],
    }


def macro_matrix_series(
    *,
    direction: Literal["input", "output"] = "input",
    workbook: str = "matrix_inputs.xlsx",
    **overrides: Any,
) -> dict[str, Any]:
    """Return an inline macro_matrix series entry for tests."""
    series: dict[str, Any] = {
        "id": "macro_matrix",
        "sheet": "Inputs",
        "data_range": "Inputs!B3:D5",
        "layout": "matrix",
        "structure": macro_matrix_structure(),
        "key": ["INDICATOR", "TIME_PERIOD"],
    }
    if direction == "input":
        series["input"] = {"setter": {"name": "set_macro_matrix"}}
    else:
        series["output"] = {"compute": {"name": "compute_macro_matrix"}}
    series.update(overrides)
    return series


def macro_matrix_bindings_document(
    *,
    direction: Literal["input", "output"] = "input",
    workbook: str = "matrix_inputs.xlsx",
    **series_overrides: Any,
) -> dict[str, Any]:
    """Return a binding manifest document for macro_matrix tests."""
    return {
        "schema_version": "1.4.0",
        "workbook": workbook,
        "series": [macro_matrix_series(direction=direction, workbook=workbook, **series_overrides)],
    }


def matrix_leaf_address_map(resolved: Any) -> dict[tuple[Any, ...], str]:
    """Map resolved composite keys to spreadsheet addresses."""
    result: dict[tuple[Any, ...], str] = {}
    for leaf in resolved["leaves"]:
        key_items = tuple(sorted(leaf["key"].items()))
        result[key_items] = leaf["address"]
    return result
