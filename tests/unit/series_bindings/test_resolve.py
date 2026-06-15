"""Unit tests for series binding coordinate resolution."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.graph import DependencyGraph
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


def test_resolve_borvelia_series_coordinates(tmp_path: Path) -> None:
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
        "layout": "series",
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


def _write_bool_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Flags")
    ws.write_boolean("A1", True)
    ws.write_boolean("B1", False)
    ws.write("A2", "TRUE")
    ws.write_number("B2", 1)
    ws.write_number("C2", 0)
    wb.close()


def test_resolve_bool_data_cell_read_auto(tmp_path: Path) -> None:
    wb_path = tmp_path / "flags.xlsx"
    _write_bool_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Flags!A2", "Flags!B2"], load_values=True)
    series = {
        "id": "bool_cells",
        "sheet": "Flags",
        "data_range": "Flags!A2:B2",
        "layout": "series",
        "setter": {"name": "set_bool_cells"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "bool",
                "bind": {"kind": "data_cell", "read": "auto"},
            },
            "dimensions": [
                {
                    "concept": "SLOT",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "bool"},
                }
            ],
        },
        "key": ["SLOT"],
    }

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True
    by_slot = {leaf["key"]["SLOT"]: leaf for leaf in resolved["leaves"]}
    assert by_slot[True]["coordinates"]["OBS_VALUE"] == "TRUE"
    assert by_slot[False]["coordinates"]["OBS_VALUE"] == 1


def test_resolve_bool_data_cell_read_bool_from_numeric(tmp_path: Path) -> None:
    wb_path = tmp_path / "flags.xlsx"
    _write_bool_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Flags!B2"], load_values=True)
    series = {
        "id": "bool_numeric",
        "sheet": "Flags",
        "data_range": "Flags!B2",
        "layout": "scalar",
        "setter": {"name": "set_bool_numeric"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "bool",
                "bind": {"kind": "data_cell", "read": "bool"},
            },
            "dimensions": [],
        },
        "key": [],
    }

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True
    assert resolved["leaves"][0]["coordinates"]["OBS_VALUE"] is True


def test_resolve_bool_constant_native_and_inferred(tmp_path: Path) -> None:
    wb_path = tmp_path / "flags.xlsx"
    _write_bool_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Flags!A2"], load_values=True)
    concept_scheme = {
        "id": "flags",
        "concepts": [{"id": "IS_ACTIVE", "dtype": "bool"}],
    }
    series = {
        "id": "bool_constant",
        "sheet": "Flags",
        "data_range": "Flags!A2",
        "layout": "scalar",
        "setter": {"name": "set_bool_constant"},
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
                    "scope": "series",
                    "bind": {"kind": "constant", "value": True},
                }
            ],
        },
        "key": ["IS_ACTIVE"],
    }

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=concept_scheme,
    )
    assert resolved["leaves"][0]["key"]["IS_ACTIVE"] is True

    series["structure"]["dimensions"][0]["bind"]["value"] = "FALSE"
    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=concept_scheme,
    )
    assert resolved["leaves"][0]["key"]["IS_ACTIVE"] is False


def _write_datetime_header_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    date_format = wb.add_format({"num_format": "yyyy-mm-dd"})
    periods = [datetime(2024, 1, 1), datetime(2024, 2, 1)]
    for col_index, period in enumerate(periods, start=1):
        ws.write_datetime(0, col_index, period, date_format)
        ws.write_number(1, col_index, float(col_index))
    wb.close()


def test_resolve_datetime_column_header_read_auto(tmp_path: Path) -> None:
    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_header_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2:C2"], load_values=True)
    series = {
        "id": "calendar_auto",
        "sheet": "Inputs",
        "data_range": "Inputs!B2:C2",
        "layout": "series",
        "setter": {"name": "set_calendar_auto"},
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
                    "bind": {"kind": "column_header", "header_row": 1, "read": "auto"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True
    periods = {leaf["key"]["TIME_PERIOD"] for leaf in resolved["leaves"]}
    assert periods == {datetime(2024, 1, 1), datetime(2024, 2, 1)}


def test_resolve_datetime_column_header_read_datetime(tmp_path: Path) -> None:
    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_header_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2"], load_values=True)
    series = {
        "id": "calendar_explicit",
        "sheet": "Inputs",
        "data_range": "Inputs!B2",
        "layout": "scalar",
        "setter": {"name": "set_calendar_explicit"},
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
                    "bind": {"kind": "column_header", "header_row": 1, "read": "datetime"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["leaves"][0]["key"]["TIME_PERIOD"] == datetime(2024, 1, 1)


def test_resolve_datetime_constant_from_iso_string(tmp_path: Path) -> None:
    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_header_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2"], load_values=True)
    concept_scheme = {
        "id": "calendar",
        "concepts": [{"id": "TIME_PERIOD", "dtype": "datetime"}],
    }
    series = {
        "id": "calendar_constant",
        "sheet": "Inputs",
        "data_range": "Inputs!B2",
        "layout": "scalar",
        "setter": {"name": "set_calendar_constant"},
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
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "2024-03-15"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=concept_scheme,
    )
    assert resolved["leaves"][0]["key"]["TIME_PERIOD"] == datetime(2024, 3, 15)


def test_resolve_datetime_bind_failure_is_reported(tmp_path: Path) -> None:
    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_header_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2"], load_values=True)
    series = {
        "id": "calendar_invalid",
        "sheet": "Inputs",
        "data_range": "Inputs!B2",
        "layout": "scalar",
        "setter": {"name": "set_calendar_invalid"},
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
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "not-a-date"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }
    concept_scheme = {
        "id": "calendar",
        "concepts": [{"id": "TIME_PERIOD", "dtype": "datetime"}],
    }

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=concept_scheme,
    )
    assert resolved["ok"] is False
    assert any(i["code"] == "bind_resolution_failed" for i in resolved["issues"])


def _scalar_series_with_series_context(
    *,
    series_context: dict[str, object],
    wb_path: Path,
) -> tuple[DependencyGraph, dict[str, Any]]:
    graph = create_dependency_graph(wb_path, ["Sheet1!B2"], load_values=True)
    series: dict[str, Any] = {
        "id": "context_series",
        "sheet": "Sheet1",
        "data_range": "Sheet1!B2",
        "layout": "scalar",
        "setter": {"name": "set_context_series"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
        },
        "series_context": series_context,
    }
    return graph, series


def test_resolve_series_context_datetime_from_iso_string(tmp_path: Path) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 1.0)
    wb.close()

    graph, series = _scalar_series_with_series_context(
        series_context={"REPORT_DATE": "2024-06-30"},
        wb_path=wb_path,
    )
    concept_scheme = {
        "id": "context",
        "concepts": [{"id": "REPORT_DATE", "dtype": "datetime"}],
    }

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=concept_scheme,
    )
    assert resolved["ok"] is True
    assert resolved["leaves"][0]["record"]["REPORT_DATE"] == datetime(2024, 6, 30)


def test_resolve_series_context_bool_from_string(tmp_path: Path) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 1.0)
    wb.close()

    graph, series = _scalar_series_with_series_context(
        series_context={"IS_ACTIVE": "TRUE"},
        wb_path=wb_path,
    )
    concept_scheme = {
        "id": "context",
        "concepts": [{"id": "IS_ACTIVE", "dtype": "bool"}],
    }

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=concept_scheme,
    )
    assert resolved["ok"] is True
    assert resolved["leaves"][0]["record"]["IS_ACTIVE"] is True


def test_resolve_series_context_without_dtype_preserves_strings(tmp_path: Path) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 1.0)
    wb.close()

    graph, series = _scalar_series_with_series_context(
        series_context={"INDICATOR": "Rec"},
        wb_path=wb_path,
    )

    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True
    assert resolved["leaves"][0]["record"]["INDICATOR"] == "Rec"


def test_resolve_series_context_invalid_datetime_fails(tmp_path: Path) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 1.0)
    wb.close()

    graph, series = _scalar_series_with_series_context(
        series_context={"REPORT_DATE": "not-a-date"},
        wb_path=wb_path,
    )
    concept_scheme = {
        "id": "context",
        "concepts": [{"id": "REPORT_DATE", "dtype": "datetime"}],
    }

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=concept_scheme,
    )
    assert resolved["ok"] is False
    assert any(i["code"] == "series_context_coercion_failed" for i in resolved["issues"])
