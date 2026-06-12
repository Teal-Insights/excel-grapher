"""Unit tests for graph-backed series binding validation."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    validate_series_bindings,
)
from excel_grapher.series_bindings.schema import validate_bindings_document

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def _write_borvelia_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year - 3))
    wb.close()


def test_expand_data_range_row() -> None:
    addresses = expand_data_range("Inputs!F5:H5")
    assert addresses == ["Inputs!F5", "Inputs!G5", "Inputs!H5"]


def test_validate_series_bindings_happy_path(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        ["Inputs!F5", "Inputs!G5", "Inputs!H5", "Inputs!I5", "Inputs!J5"],
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True
    assert not any(i["level"] == "error" for i in report["issues"])


def test_validate_allows_partial_graph_leaf_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!F5"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True
    codes = {i["code"] for i in report["issues"]}
    assert "missing_graph_node" not in codes


def test_validate_allows_zero_graph_leaf_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!A2"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True
    assert not any(i["level"] == "error" for i in report["issues"])


def test_validate_reports_non_leaf_formula_cell(tmp_path: Path) -> None:
    wb_path = tmp_path / "model.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    ws.write_formula("F5", "=G5")
    ws.write_number("G5", 100)
    ws.write_number("H5", 2)
    ws.write_number("I5", 3)
    ws.write_number("J5", 4)
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
    wb.close()

    graph = create_dependency_graph(
        wb_path,
        ["Inputs!F5", "Inputs!G5", "Inputs!H5", "Inputs!I5", "Inputs!J5"],
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True
    assert not any(i["code"] == "not_a_leaf" for i in report["issues"])


def test_validate_warns_when_measure_dtype_and_read_disagree(tmp_path: Path) -> None:
    wb_path = tmp_path / "flags.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Flags")
    ws.write_boolean("B2", True)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Flags!B2"], load_values=True)
    bindings = validate_bindings_document(
        {
            "schema_version": "1.4.0",
            "workbook": "flags.xlsx",
            "series": [
                {
                    "id": "bool_mismatch",
                    "sheet": "Flags",
                    "data_range": "Flags!B2",
                    "layout": "scalar",
                    "setter": {"name": "set_bool_mismatch"},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "bool",
                            "bind": {"kind": "data_cell", "read": "int"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                }
            ],
        }
    )
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True
    warnings = [issue for issue in report["issues"] if issue["code"] == "dtype_read_mismatch"]
    assert len(warnings) == 1
    assert "measure dtype 'bool'" in warnings[0]["message"]


def test_validate_warns_when_dimension_dtype_and_read_disagree(tmp_path: Path) -> None:
    wb_path = tmp_path / "flags.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Flags")
    ws.write_boolean("B1", True)
    ws.write_boolean("B2", False)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Flags!B2"], load_values=True)
    bindings = validate_bindings_document(
        {
            "schema_version": "1.4.0",
            "workbook": "flags.xlsx",
            "concept_scheme": {
                "id": "flags",
                "concepts": [{"id": "IS_ACTIVE", "dtype": "bool"}],
            },
            "series": [
                {
                    "id": "bool_mismatch",
                    "sheet": "Flags",
                    "data_range": "Flags!B2",
                    "layout": "scalar",
                    "setter": {"name": "set_bool_mismatch"},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "bool",
                            "bind": {"kind": "data_cell", "read": "bool"},
                        },
                        "dimensions": [
                            {
                                "concept": "IS_ACTIVE",
                                "role": "key",
                                "scope": "cell",
                                "bind": {
                                    "kind": "column_header",
                                    "header_row": 1,
                                    "read": "int",
                                },
                            }
                        ],
                    },
                    "key": ["IS_ACTIVE"],
                }
            ],
        }
    )
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True
    warnings = [issue for issue in report["issues"] if issue["code"] == "dtype_read_mismatch"]
    assert len(warnings) == 1
    assert "dimension 'IS_ACTIVE'" in warnings[0]["message"]
