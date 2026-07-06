"""Resolution tests for grouped-row matrix geometry (schema 1.5.0).

Covers `skip`/`include` label-source restriction, `fill` propagation,
`missing` policy, the `value_map` bind kind, and series-level `exclude_rows`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)
from tests.fixtures.series_bindings.grouped_matrix_helpers import (
    MATRIX_GROUPED_ROWS_BINDINGS,
    expected_grouped_matrix_keys,
    grouped_matrix_series,
    write_grouped_matrix_workbook,
)


def _grouped_graph(tmp_path: Path) -> tuple[Path, Any]:
    wb_path = tmp_path / "grouped_inputs.xlsx"
    write_grouped_matrix_workbook(wb_path)
    targets = expand_data_range("Inputs!C2:D8")
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    return wb_path, graph


def test_resolve_grouped_matrix_fixture(tmp_path: Path) -> None:
    wb_path, graph = _grouped_graph(tmp_path)
    bindings = load_series_bindings(MATRIX_GROUPED_ROWS_BINDINGS)
    series = bindings["series"][0]

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True, resolved["issues"]
    assert resolved["requires_address"] is False
    assert len(resolved["leaves"]) == 8

    by_key = {
        (
            leaf["key"]["SCENARIO"],
            leaf["key"]["SHOCK_TYPE"],
            leaf["key"]["TIME_PERIOD"],
        ): leaf
        for leaf in resolved["leaves"]
    }
    assert set(by_key) == expected_grouped_matrix_keys()

    paris_revenue_2024 = by_key[("Paris", "Revenue", 2024)]
    assert paris_revenue_2024["address"] == "Inputs!C3"
    assert paris_revenue_2024["coordinates"]["OBS_VALUE"] == 1.1

    moderate_pe_2025 = by_key[("Moderate", "Primary expenditure", 2025)]
    assert moderate_pe_2025["address"] == "Inputs!D8"
    assert moderate_pe_2025["coordinates"]["OBS_VALUE"] == 4.2


def test_exclude_rows_drops_referenced_blank_rows(tmp_path: Path) -> None:
    """Blank header/separator rows are graph leaves (SUM over the block) but excluded."""
    wb_path, graph = _grouped_graph(tmp_path)
    assert "Inputs!C2" in graph
    assert "Inputs!C5" in graph

    series = grouped_matrix_series()
    resolved = resolve_series_binding(graph, wb_path, series)
    addresses = {leaf["address"] for leaf in resolved["leaves"]}
    for excluded in ("Inputs!C2", "Inputs!D2", "Inputs!C5", "Inputs!D5", "Inputs!C6", "Inputs!D6"):
        assert excluded not in addresses
    assert len(addresses) == 8


def test_missing_label_errors_by_default_without_exclude_rows(tmp_path: Path) -> None:
    """Without exclude_rows, unlabeled rows fail loudly instead of mis-keying."""
    wb_path, graph = _grouped_graph(tmp_path)
    series = grouped_matrix_series()
    del series["exclude_rows"]

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is False
    codes = {issue["code"] for issue in resolved["issues"]}
    assert "bind_resolution_failed" in codes


def test_missing_null_yields_none_coordinate(tmp_path: Path) -> None:
    wb_path = tmp_path / "sparse.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write(0, 2, 2024)
    ws.write("A2", "Grouped")
    ws.write_number("C2", 1.0)
    ws.write_number("C3", 2.0)  # A3 has no label
    wb.close()
    graph = create_dependency_graph(wb_path, ["Inputs!C2", "Inputs!C3"], load_values=True)

    series = {
        "id": "sparse",
        "sheet": "Inputs",
        "data_range": "Inputs!C2:C3",
        "layout": "matrix",
        "input": {"setter": {"name": "set_sparse"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
            "dimensions": [
                {
                    "concept": "GROUP",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "row_label",
                        "label_column": "A",
                        "missing": "null",
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
        },
        "key": ["GROUP", "TIME_PERIOD"],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True, resolved["issues"]
    by_address = {leaf["address"]: leaf for leaf in resolved["leaves"]}
    assert by_address["Inputs!C2"]["key"]["GROUP"] == "Grouped"
    assert by_address["Inputs!C3"]["key"]["GROUP"] is None


def test_fill_down_covers_merged_style_sparse_group_column(tmp_path: Path) -> None:
    """Sparse (or merged) group labels in column A fill down over their band."""
    wb_path = tmp_path / "sparse_groups.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write(0, 2, 2024)
    rows = [
        ("Paris", "Revenue", 1.0),
        (None, "Primary expenditure", 2.0),
        ("Moderate", "Revenue", 3.0),
        (None, "Primary expenditure", 4.0),
    ]
    for offset, (scenario, shock, value) in enumerate(rows):
        if scenario is not None:
            ws.write(1 + offset, 0, scenario)
        ws.write(1 + offset, 1, shock)
        ws.write_number(1 + offset, 2, value)
    wb.close()
    targets = expand_data_range("Inputs!C2:C5")
    graph = create_dependency_graph(wb_path, targets, load_values=True)

    series = {
        "id": "sparse_groups",
        "sheet": "Inputs",
        "data_range": "Inputs!C2:C5",
        "layout": "matrix",
        "input": {"setter": {"name": "set_sparse_groups"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
            "dimensions": [
                {
                    "concept": "SCENARIO",
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
                    "concept": "SHOCK_TYPE",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "B", "read": "string"},
                },
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["SCENARIO", "SHOCK_TYPE", "TIME_PERIOD"],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True, resolved["issues"]
    keys = {(leaf["key"]["SCENARIO"], leaf["key"]["SHOCK_TYPE"]) for leaf in resolved["leaves"]}
    assert keys == {
        ("Paris", "Revenue"),
        ("Paris", "Primary expenditure"),
        ("Moderate", "Revenue"),
        ("Moderate", "Primary expenditure"),
    }


def test_value_map_binds_rows_without_sheet_labels(tmp_path: Path) -> None:
    """value_map keys rows by manifest values when labels are absent from the sheet."""
    wb_path = tmp_path / "no_labels.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write(0, 2, 2024)
    for offset, value in enumerate([1.0, 2.0, 3.0, 4.0]):
        ws.write_number(1 + offset, 2, value)
    wb.close()
    targets = expand_data_range("Inputs!C2:C5")
    graph = create_dependency_graph(wb_path, targets, load_values=True)

    series = {
        "id": "no_labels",
        "sheet": "Inputs",
        "data_range": "Inputs!C2:C5",
        "layout": "matrix",
        "input": {"setter": {"name": "set_no_labels"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
            "dimensions": [
                {
                    "concept": "SCENARIO",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "value_map",
                        "values": {"Paris": "2:3", "Moderate": [4, 5]},
                    },
                },
                {
                    "concept": "SHOCK_TYPE",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "value_map",
                        "values": {"Revenue": [2, 4], "Primary expenditure": [3, 5]},
                    },
                },
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["SCENARIO", "SHOCK_TYPE", "TIME_PERIOD"],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True, resolved["issues"]
    keys = {(leaf["key"]["SCENARIO"], leaf["key"]["SHOCK_TYPE"]) for leaf in resolved["leaves"]}
    assert keys == {
        ("Paris", "Revenue"),
        ("Paris", "Primary expenditure"),
        ("Moderate", "Revenue"),
        ("Moderate", "Primary expenditure"),
    }


def test_value_map_uncovered_row_errors_by_default(tmp_path: Path) -> None:
    wb_path = tmp_path / "uncovered.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write(0, 2, 2024)
    ws.write_number("C2", 1.0)
    ws.write_number("C3", 2.0)
    wb.close()
    graph = create_dependency_graph(wb_path, ["Inputs!C2", "Inputs!C3"], load_values=True)

    series = {
        "id": "uncovered",
        "sheet": "Inputs",
        "data_range": "Inputs!C2:C3",
        "layout": "matrix",
        "input": {"setter": {"name": "set_uncovered"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
            "dimensions": [
                {
                    "concept": "SCENARIO",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "value_map", "values": {"Paris": 2}},
                },
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is False
    codes = {issue["code"] for issue in resolved["issues"]}
    assert "bind_resolution_failed" in codes


def test_value_map_column_axis(tmp_path: Path) -> None:
    """Column-letter specs key data columns instead of rows."""
    wb_path = tmp_path / "col_map.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Revenue")
    ws.write_number("C2", 1.0)
    ws.write_number("D2", 2.0)
    wb.close()
    graph = create_dependency_graph(wb_path, ["Inputs!C2", "Inputs!D2"], load_values=True)

    series = {
        "id": "col_map",
        "sheet": "Inputs",
        "data_range": "Inputs!C2:D2",
        "layout": "series",
        "input": {"setter": {"name": "set_col_map"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
            "dimensions": [
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "value_map", "values": {2024: "C", 2025: "D"}},
                },
            ],
        },
        "key": ["TIME_PERIOD"],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True, resolved["issues"]
    by_key = {leaf["key"]["TIME_PERIOD"]: leaf["address"] for leaf in resolved["leaves"]}
    assert by_key == {2024: "Inputs!C2", 2025: "Inputs!D2"}


def test_column_header_fill_right_covers_merged_style_headers(tmp_path: Path) -> None:
    """A header spanning several columns (merged anchor) fills rightward."""
    wb_path = tmp_path / "wide_header.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write("C1", "Baseline")  # D1 empty: covered cell of a merged header
    ws.write("A2", "Revenue")
    ws.write_number("C2", 1.0)
    ws.write_number("D2", 2.0)
    wb.close()
    graph = create_dependency_graph(wb_path, ["Inputs!C2", "Inputs!D2"], load_values=True)

    series = {
        "id": "wide_header",
        "sheet": "Inputs",
        "data_range": "Inputs!C2:D2",
        "layout": "series",
        "input": {"setter": {"name": "set_wide_header"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
            "dimensions": [
                {
                    "concept": "SCENARIO",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "column_header",
                        "header_row": 1,
                        "fill": True,
                        "read": "string",
                    },
                },
                {
                    "concept": "OFFSET",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "value_map",
                        "values": {1: "C", 2: "D"},
                    },
                },
            ],
        },
        "key": ["SCENARIO", "OFFSET"],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True, resolved["issues"]
    scenarios = {leaf["key"]["SCENARIO"] for leaf in resolved["leaves"]}
    assert scenarios == {"Baseline"}
