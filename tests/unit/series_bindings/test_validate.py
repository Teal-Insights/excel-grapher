"""Unit tests for graph-backed series binding validation."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import fastpyxl
import xlsxwriter

from excel_grapher.grapher import DependencyGraph, Node, create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    validate_series_bindings,
)
from excel_grapher.series_bindings.schema import validate_bindings_document
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


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
    assert report["ok"] is False
    assert any(i["code"] == "non_leaf_input_overlap" for i in report["issues"])


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


def test_validate_warns_when_attribute_bind_read_disagrees_with_concept_dtype(
    tmp_path: Path,
) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 1.0)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Sheet1!B2"], load_values=True)
    bindings = validate_bindings_document(
        {
            "schema_version": "1.4.0",
            "workbook": "scalar.xlsx",
            "concept_scheme": {
                "id": "attrs",
                "concepts": [{"id": "REPORT_DATE", "dtype": "datetime"}],
            },
            "series": [
                {
                    "id": "attr_mismatch",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!B2",
                    "layout": "scalar",
                    "setter": {"name": "set_attr_mismatch"},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                        "attributes": [
                            {
                                "concept": "REPORT_DATE",
                                "role": "attribute",
                                "bind": {
                                    "kind": "cell",
                                    "address": "Sheet1!A1",
                                    "read": "string",
                                },
                            }
                        ],
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
    assert "attribute 'REPORT_DATE'" in warnings[0]["message"]


def test_validate_skips_attribute_value_shorthand_for_dtype_read_check(
    tmp_path: Path,
) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 1.0)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Sheet1!B2"], load_values=True)
    bindings = validate_bindings_document(
        {
            "schema_version": "1.4.0",
            "workbook": "scalar.xlsx",
            "concept_scheme": {
                "id": "attrs",
                "concepts": [{"id": "REPORT_DATE", "dtype": "datetime"}],
            },
            "series": [
                {
                    "id": "attr_constant",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!B2",
                    "layout": "scalar",
                    "setter": {"name": "set_attr_constant"},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                        "attributes": [
                            {
                                "concept": "REPORT_DATE",
                                "role": "attribute",
                                "value": "2024-06-30",
                            }
                        ],
                    },
                    "key": [],
                }
            ],
        }
    )
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True
    assert not any(issue["code"] == "dtype_read_mismatch" for issue in report["issues"])


def test_validate_matrix_layout_intent_requires_two_dimensions(tmp_path: Path) -> None:
    from tests.fixtures.series_bindings.matrix_helpers import (
        macro_matrix_bindings_document,
        write_matrix_explicit_workbook,
    )

    wb_path = tmp_path / "matrix_inputs.xlsx"
    write_matrix_explicit_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!B3:D5"),
        load_values=True,
    )
    bindings = validate_bindings_document(macro_matrix_bindings_document())
    bindings["series"][0]["structure"]["dimensions"] = bindings["series"][0]["structure"][
        "dimensions"
    ][:1]
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is False
    assert any(issue["code"] == "layout_constraint_violation" for issue in report["issues"])


def test_validate_series_layout_requires_cell_scoped_dimension(tmp_path: Path) -> None:
    wb_path = tmp_path / "series_layout.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write_number("B2", 1.0)
    wb.close()
    graph = create_dependency_graph(wb_path, ["Inputs!B2"], load_values=True)
    bindings: WorkbookSeriesBindings = {
        "schema_version": "1.4.0",
        "series": [
            {
                "id": "series_only_scope",
                "sheet": "Inputs",
                "data_range": "Inputs!B2",
                "layout": "series",
                "input": {"setter": {"name": "set_series_only_scope"}},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "bind": {"kind": "data_cell", "read": "float"},
                    },
                    "dimensions": [
                        {
                            "concept": "COUNTRY",
                            "role": "key",
                            "scope": "series",
                            "bind": {
                                "kind": "constant",
                                "value": "Borvelia",
                            },
                        }
                    ],
                },
                "key": ["COUNTRY"],
            }
        ],
    }
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is False
    assert any(
        issue["code"] == "layout_constraint_violation" and "cell-scoped" in issue["message"]
        for issue in report["issues"]
    )


def _matrix_row_series(series_id: str, row: int) -> dict:
    return {
        "id": series_id,
        "sheet": "S",
        "layout": "matrix",
        "data_range": f"S!B{row}:C{row}",
        "key": ["INDICATOR", "TIME_PERIOD"],
        "input": {"mode": "leaf"},
        "structure": {
            "dimensions": [
                {
                    "id": "INDICATOR",
                    "concept": "INDICATOR",
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
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
        },
        "validation": {"require_unique_key": True},
    }


def _write_matrix_demo_workbook(path: Path) -> DependencyGraph:
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = "label"
    ws["B1"] = 2020
    ws["C1"] = 2021
    for row, label, b, c in [(2, "x", 1.0, 2.0), (3, "y", 3.0, 4.0), (4, "z", 5.0, 6.0)]:
        ws[f"A{row}"] = label
        ws[f"B{row}"] = b
        ws[f"C{row}"] = c
    wb.save(path)

    graph = DependencyGraph()
    for col, row, val in [
        ("B", 2, 1.0),
        ("C", 2, 2.0),
        ("B", 3, 3.0),
        ("C", 3, 4.0),
        ("B", 4, 5.0),
        ("C", 4, 6.0),
    ]:
        graph.add_node(
            Node(
                sheet="S",
                column=col,
                row=row,
                formula=None,
                normalized_formula=None,
                value=val,
                is_leaf=True,
            )
        )
    return graph


def test_validate_series_bindings_loads_workbook_once(tmp_path: Path) -> None:
    """Validate should share one workbook reader across series (issue #416)."""
    wb_path = tmp_path / "demo.xlsx"
    graph = _write_matrix_demo_workbook(wb_path)
    series_list = [_matrix_row_series(f"s{row}", row) for row in (2, 3, 4)]
    bindings: WorkbookSeriesBindings = {
        "schema_version": "1.4.0",
        "workbook": wb_path.name,
        "series": series_list,
    }

    calls: list[dict] = []
    real_load = fastpyxl.load_workbook

    def wrapped(*args, **kwargs):
        calls.append(dict(kwargs))
        return real_load(*args, **kwargs)

    with patch("fastpyxl.load_workbook", side_effect=wrapped):
        report = validate_series_bindings(graph, bindings, workbook=wb_path)

    assert report["ok"] is True
    assert len(calls) == 1
    assert calls[0].get("data_only") is True
    assert calls[0].get("read_only") is True
