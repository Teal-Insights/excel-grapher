"""Omit `compute_all` when output series bindings cover every export target (#457)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document

BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.10.0",
    "concept_scheme": {
        "id": "m",
        "concepts": [
            {"id": "OBS_VALUE", "name": "v", "dtype": "number"},
            {"id": "TIME_PERIOD", "name": "t", "dtype": "int"},
        ],
    },
    "series": [
        {
            "id": "out_series",
            "sheet": "Out",
            "data_range": "Out!B1:D1",
            "layout": "series",
            "output": {"compute": {"name": "compute_out_series"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "TIME_PERIOD",
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 2, "read": "int"},
                    }
                ],
            },
            "key": ["TIME_PERIOD"],
        }
    ],
}


def _write_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    dash = wb.add_worksheet("Dash")
    dash.write("B1", "Alpha")
    data = wb.add_worksheet("Data")
    data.write("A1", "Alpha")
    data.write_number("B1", 10.0)
    data.write("A2", "Beta")
    data.write_number("B2", 20.0)
    out = wb.add_worksheet("Out")
    for col in range(1, 4):  # B, C, D
        out.write_formula(0, col, "=VLOOKUP(Dash!B1,Data!A1:B2,2,FALSE)")
        out.write_number(1, col, 2028 + col + 1)
    wb.close()


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "m.xlsx"
    _write_workbook(path)
    return path


def _targets(workbook: Path) -> list[str]:
    return expand_data_range("Out!B1:D1", workbook=workbook)


def _bindings() -> Any:
    return validate_bindings_document(dict(BINDINGS_DOCUMENT))


def test_generate_omits_compute_all_when_output_bindings_cover_targets(workbook: Path) -> None:
    targets = _targets(workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)
    source = CodeGenerator(graph).generate(
        targets,
        series_bindings=_bindings(),
        bindings_workbook=workbook,
    )

    assert "def compute_out_series(" in source
    assert "def compute_all(" not in source
    assert "TARGETS = {" not in source
