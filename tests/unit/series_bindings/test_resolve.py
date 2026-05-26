"""Unit tests for series binding coordinate resolution."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
    resolve_series_bindings,
)

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


def test_resolve_borvelia_row_series_coordinates(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    targets = expand_data_range("Inputs!F5:J5")
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True
    assert resolved["requires_address"] is False
    assert len(resolved["leaves"]) == 5

    by_period = {leaf["key"]["TIME_PERIOD"]: leaf for leaf in resolved["leaves"]}
    assert set(by_period) == {1, 2, 3, 4, 5}
    leaf = by_period[3]
    assert leaf["address"] == "Inputs!H5"
    assert leaf["coordinates"]["REF_AREA"] == "Borvelia"
    assert leaf["coordinates"]["INDICATOR"] == "Primary balance"
    assert leaf["coordinates"]["TIME_PERIOD"] == 3
    assert leaf["coordinates"]["OBS_VALUE"] == 0.0
    assert leaf["record"]["OBS_VALUE"] == 0.0
    assert leaf["record"]["TIME_PERIOD"] == 3
    assert leaf["record"]["REF_AREA"] == "Borvelia"
    assert leaf["record"]["UNIT_MEASURE"] == "PC_GDP"
    assert leaf["record"]["INDICATOR"] == "Primary balance (% of GDP)"


def test_resolve_filters_series_to_graph_leaves(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!H5"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]

    resolved = resolve_series_binding(graph, wb_path, series)

    assert resolved["ok"] is True
    assert [leaf["address"] for leaf in resolved["leaves"]] == ["Inputs!H5"]
    assert resolved["leaves"][0]["key"] == {"TIME_PERIOD": 3}


def test_resolve_zero_graph_leaf_overlap_returns_no_leaves(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!A2"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]

    resolved = resolve_series_binding(graph, wb_path, series)

    assert resolved["ok"] is True
    assert resolved["leaves"] == []


def test_resolve_scalar_binding(tmp_path: Path) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B25", 0.05)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Sheet1!B25"], load_values=True)
    series = {
        "id": "scalar_threshold_p",
        "sheet": "Sheet1",
        "data_range": "Sheet1!B25",
        "layout": "scalar",
        "setter": {"name": "set_scalar_threshold_p"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "concept": "INPUT_NAME",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "threshold_p_value"},
                }
            ],
        },
        "key": ["INPUT_NAME"],
        "series_context": {"INPUT_NAME": "threshold_p_value"},
    }

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True
    assert len(resolved["leaves"]) == 1
    leaf = resolved["leaves"][0]
    assert leaf["key"] == {"INPUT_NAME": "threshold_p_value"}
    assert leaf["record"]["OBS_VALUE"] == pytest.approx(0.05)


def test_resolve_duplicate_key_sets_requires_address(tmp_path: Path) -> None:
    wb_path = tmp_path / "dup.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("C1", 1)
    ws.write_number("D1", 1)
    ws.write_number("C2", 10)
    ws.write_number("D2", 20)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Sheet1!C2", "Sheet1!D2"], load_values=True)
    series = {
        "id": "dup_headers",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2:D2",
        "layout": "row_series",
        "setter": {"name": "set_dup_headers"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
        "validation": {"require_unique_key": True},
    }

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["requires_address"] is True
    assert any(i["code"] == "duplicate_key" for i in resolved["issues"])


def test_resolve_series_bindings_workbook(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    report = resolve_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True
    assert len(report["series"]) == 1
    assert report["series"][0]["requires_address"] is False
