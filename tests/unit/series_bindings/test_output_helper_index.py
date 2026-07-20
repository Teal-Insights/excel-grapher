"""Unit tests for output address → parameterized helper coverage index."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import resolve_series_binding
from excel_grapher.series_bindings.output_helper_index import (
    OutputHelperSpec,
    build_output_helper_index,
    format_output_helper_call_form,
    resolve_output_helper_ref,
)
from excel_grapher.series_bindings.types import WorkbookSeriesBindings


def _write_series_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 2, 2020)
    ws.write_number(0, 3, 2021)
    ws.write_number("B2", 5.0)
    ws.write_formula("C2", "=B2*2")
    ws.write_formula("D2", "=B2*3")
    wb.close()


def _series(*, helper: dict[str, Any] | None = None) -> dict[str, Any]:
    compute: dict[str, Any] = {"name": "compute_scaled_output"}
    if helper is not None:
        compute["helper"] = helper
    return {
        "id": "scaled_output",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2:D2",
        "layout": "series",
        "output": {"compute": compute},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
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
    }


def test_format_output_helper_call_form_from_dims() -> None:
    assert (
        format_output_helper_call_form(
            "scaled_output_hot",
            dims=["TIME_PERIOD"],
            record_expr="static_record",
        )
        == "scaled_output_hot(ctx, time_period=static_record['TIME_PERIOD'])"
    )
    assert (
        format_output_helper_call_form(
            "macro_matrix_helper",
            dims=["INDICATOR", "TIME_PERIOD"],
            record_expr="static_record",
        )
        == "macro_matrix_helper(ctx, indicator=static_record['INDICATOR'], "
        "time_period=static_record['TIME_PERIOD'])"
    )


def test_build_index_from_bindings_helper_defaults_dims_to_key(tmp_path: Path) -> None:
    wb_path = tmp_path / "scaled.xlsx"
    _write_series_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2", "Sheet1!D2"], load_values=True)
    series = _series(helper={"name": "scaled_output_hot"})
    bindings = cast(
        WorkbookSeriesBindings,
        {
            "schema_version": "1.10.0",
            "workbook": "scaled.xlsx",
            "series": [series],
        },
    )
    index = build_output_helper_index(graph, bindings, workbook=wb_path)
    assert set(index["leaves"]) == {"Sheet1!C2", "Sheet1!D2"}
    entry = index["leaves"]["Sheet1!C2"]
    assert entry["helper"] == "scaled_output_hot"
    assert entry["dims"] == ["TIME_PERIOD"]
    assert entry["series_id"] == "scaled_output"
    assert entry["call_form"] == (
        "scaled_output_hot(ctx, time_period=static_record['TIME_PERIOD'])"
    )


def test_build_index_respects_explicit_dims_and_address_overlay(tmp_path: Path) -> None:
    wb_path = tmp_path / "scaled.xlsx"
    _write_series_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2", "Sheet1!D2"], load_values=True)
    series = _series(helper={"name": "scaled_output_hot", "dims": ["TIME_PERIOD"]})
    bindings = cast(
        WorkbookSeriesBindings,
        {
            "schema_version": "1.10.0",
            "workbook": "scaled.xlsx",
            "series": [series],
        },
    )
    overlay: dict[str, OutputHelperSpec] = {
        "Sheet1!D2": {"helper": "other_helper", "dims": ["TIME_PERIOD"]},
    }
    index = build_output_helper_index(
        graph,
        bindings,
        workbook=wb_path,
        address_helpers=overlay,
    )
    assert index["leaves"]["Sheet1!C2"]["helper"] == "scaled_output_hot"
    assert index["leaves"]["Sheet1!D2"]["helper"] == "other_helper"


def test_resolve_output_helper_ref_falls_back_to_xl_cell(tmp_path: Path) -> None:
    wb_path = tmp_path / "scaled.xlsx"
    _write_series_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2", "Sheet1!D2"], load_values=True)
    series = _series(helper={"name": "scaled_output_hot"})
    bindings = cast(
        WorkbookSeriesBindings,
        {
            "schema_version": "1.10.0",
            "workbook": "scaled.xlsx",
            "series": [series],
        },
    )
    index = build_output_helper_index(graph, bindings, workbook=wb_path)
    covered = resolve_output_helper_ref("Sheet1!C2", index=index)
    assert covered["mode"] == "helper"
    assert covered["helper"] == "scaled_output_hot"
    unbound = resolve_output_helper_ref("Sheet1!Z99", index=index)
    assert unbound["mode"] == "xl_cell"
    assert unbound["reason"] == "unbound"
    assert unbound["call_form"] == "xl_cell(ctx, 'Sheet1!Z99')"


def test_resolve_series_binding_still_works_with_helper_metadata(tmp_path: Path) -> None:
    wb_path = tmp_path / "scaled.xlsx"
    _write_series_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2", "Sheet1!D2"], load_values=True)
    series = _series(helper={"name": "scaled_output_hot", "dims": ["TIME_PERIOD"]})
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    assert resolved["ok"] is True
    assert len(resolved["leaves"]) == 2
