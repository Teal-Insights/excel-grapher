"""Range readers honour exclude_rows / exclude_columns in the reader index.

Keyed readers already filter via resolve. Binding-aligned `read_*_range`
entries in the reader index use the narrowed contiguous rectangle after
exclusions, not the original `data_range`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    validate_bindings_document,
)
from excel_grapher.series_bindings.reader_index import build_reader_index


def _interleaved_matrix_doc() -> dict[str, Any]:
    """Two matrix series sharing Risks!B2:D5 with complementary exclude_rows."""
    structure = {
        "measure": {
            "concept": "OBS_VALUE",
            "dtype": "float",
            "bind": {"kind": "data_cell", "read": "float"},
        },
        "dimensions": [
            {
                "id": "BAND",
                "concept": "BAND",
                "dtype": "string",
                "role": "key",
                "scope": "cell",
                "bind": {
                    "kind": "row_label",
                    "label_column": "A",
                    "fill": True,
                    "read": "string",
                },
            },
            {
                "id": "TIME_PERIOD",
                "concept": "TIME_PERIOD",
                "dtype": "int",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
            },
        ],
    }
    return {
        "schema_version": "1.12.0",
        "concept_scheme": {
            "id": "mcve",
            "concepts": [
                {"id": "OBS_VALUE", "name": "Observation value", "dtype": "number"},
                {"id": "TIME_PERIOD", "name": "Time period", "dtype": "int"},
                {"id": "BAND", "name": "Band", "dtype": "string"},
            ],
        },
        "series": [
            {
                "id": "revenue_shocks",
                "sheet": "Risks",
                "data_range": "Risks!B2:D5",
                "layout": "matrix",
                "exclude_rows": [3, 5],
                "input": {"setter": {"name": "set_revenue_shocks"}},
                "structure": structure,
                "key": ["BAND", "TIME_PERIOD"],
            },
            {
                "id": "expenditure_shocks",
                "sheet": "Risks",
                "data_range": "Risks!B2:D5",
                "layout": "matrix",
                "exclude_rows": [2, 4],
                "input": {"setter": {"name": "set_expenditure_shocks"}},
                "structure": structure,
                "key": ["BAND", "TIME_PERIOD"],
            },
        ],
    }


def _write_interleaved_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Risks")
    for col, year in enumerate([2030, 2031, 2032], start=1):
        ws.write(0, col, year)
    for row, label in enumerate(["rev", "exp", "rev", "exp"], start=1):
        ws.write(row, 0, label)
        for col in range(1, 4):
            ws.write_number(row, col, float(row * 10 + col))
    ws.write_formula("E1", "=SUM(B2:D5)")
    wb.close()


def _write_edge_trim_workbook(path: Path) -> None:
    """Contiguous block B2:D5 where excluding edge rows/cols yields one rectangle."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Demo")
    for col, year in enumerate([2028, 2029, 2030], start=1):
        ws.write(0, col, year)
    for row, label in enumerate(["A", "B", "C", "D"], start=1):
        ws.write(row, 0, label)
        for col in range(1, 4):
            ws.write_number(row, col, float(row * 10 + col))
    ws.write_formula("E1", "=SUM(B2:D5)")
    wb.close()


def _matrix_series(
    *,
    series_id: str = "demo",
    data_range: str = "Demo!B2:D5",
    exclude_rows: list[Any] | None = None,
    exclude_columns: list[Any] | None = None,
) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": series_id,
        "sheet": "Demo",
        "data_range": data_range,
        "layout": "matrix",
        "input": {"setter": {"name": f"set_{series_id}"}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "BAND",
                    "concept": "BAND",
                    "dtype": "string",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "row_label",
                        "label_column": "A",
                        "read": "string",
                    },
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "dtype": "int",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["BAND", "TIME_PERIOD"],
    }
    if exclude_rows is not None:
        series["exclude_rows"] = exclude_rows
    if exclude_columns is not None:
        series["exclude_columns"] = exclude_columns
    return {
        "schema_version": "1.12.0",
        "series": [series],
    }


def test_reader_index_keys_narrowed_contiguous_range(tmp_path: Path) -> None:
    wb_path = tmp_path / "trim.xlsx"
    _write_edge_trim_workbook(wb_path)
    bindings = validate_bindings_document(
        _matrix_series(exclude_rows=[5], exclude_columns=["D"], data_range="Demo!B2:D5")
    )
    graph = create_dependency_graph(wb_path, expand_data_range("Demo!B2:D5"), load_values=True)
    index = build_reader_index(graph, bindings, workbook=wb_path)

    assert "Demo!B2:D5" not in index["ranges"]
    assert "Demo!B2:C4" in index["ranges"]
    assert index["ranges"]["Demo!B2:C4"]["reader"] == "read_demo_range"
    assert "Demo!B2:D5" not in index["ambiguous"]
