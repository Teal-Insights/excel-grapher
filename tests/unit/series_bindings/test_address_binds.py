"""Schema and resolve tests for `row_index` / `column_letter` bind kinds (#949)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    SeriesBindingsSchemaError,
    expand_data_range,
    resolve_series_binding,
    validate_bindings_document,
)
from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION


def _pin_series(
    *,
    data_range: str | list[str],
    dimensions: list[dict[str, Any]],
    key: list[str],
    series_id: str = "pins",
) -> dict[str, Any]:
    return {
        "id": series_id,
        "sheet": "Engine",
        "data_range": data_range,
        "layout": "series",
        "constant": {},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": dimensions,
        },
        "key": key,
    }


def _row_dim() -> dict[str, Any]:
    return {
        "id": "ROW",
        "concept": "ROW",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_index", "read": "int"},
    }


def _col_dim() -> dict[str, Any]:
    return {
        "id": "COL",
        "concept": "COL",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_letter", "read": "string"},
    }


def _doc(series: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": CURRENT_SCHEMA_VERSION, "series": [series]}


def test_schema_accepts_row_index_and_column_letter_binds() -> None:
    series = _pin_series(
        data_range=["Engine!B2:C2", "Engine!B4"],
        dimensions=[_row_dim(), _col_dim()],
        key=["ROW", "COL"],
    )
    bindings = validate_bindings_document(_doc(series))
    binds = [dim["bind"] for dim in bindings["series"][0]["structure"]["dimensions"]]
    assert binds == [
        {"kind": "row_index", "read": "int"},
        {"kind": "column_letter", "read": "string"},
    ]


@pytest.mark.parametrize(
    "bind",
    [
        {"kind": "row_index", "values": {2: 2}},
        {"kind": "column_letter", "address": "Engine!D18"},
    ],
)
def test_schema_rejects_address_binds_with_unknown_properties(bind: dict[str, Any]) -> None:
    series = _pin_series(
        data_range="Engine!D18",
        dimensions=[
            {
                "id": "AXIS",
                "concept": "AXIS",
                "role": "key",
                "scope": "cell",
                "bind": bind,
            }
        ],
        key=["AXIS"],
    )
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(_doc(series))


def test_resolve_row_index_and_column_letter_from_data_cell(tmp_path: Path) -> None:
    wb_path = tmp_path / "d18.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("D18", 7.0)
    wb.close()

    series = validate_bindings_document(
        _doc(
            _pin_series(
                data_range="Engine!D18",
                dimensions=[_row_dim(), _col_dim()],
                key=["ROW", "COL"],
            )
        )
    )["series"][0]
    graph = create_dependency_graph(wb_path, ["Engine!D18"], load_values=True)
    resolved = resolve_series_binding(graph, wb_path, series, direction="constant")
    assert resolved["ok"] is True, resolved["issues"]
    assert len(resolved["leaves"]) == 1
    leaf = resolved["leaves"][0]
    assert leaf["address"] == "Engine!D18"
    assert leaf["key"] == {"ROW": 18, "COL": "D"}


def test_resolve_grab_bag_needs_both_address_axes(tmp_path: Path) -> None:
    wb_path = tmp_path / "grab_bag.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("B2", 1.0)
    ws.write_number("C2", 2.0)
    ws.write_number("B4", 3.0)
    wb.close()

    targets = expand_data_range("Engine!B2:C2") + expand_data_range("Engine!B4")
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    series = validate_bindings_document(
        _doc(
            _pin_series(
                data_range=["Engine!B2:C2", "Engine!B4"],
                dimensions=[_row_dim(), _col_dim()],
                key=["ROW", "COL"],
            )
        )
    )["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series, direction="constant")
    assert resolved["ok"] is True, resolved["issues"]
    by_key = {
        (leaf["key"]["ROW"], leaf["key"]["COL"]): leaf["address"] for leaf in resolved["leaves"]
    }
    assert by_key == {(2, "B"): "Engine!B2", (2, "C"): "Engine!C2", (4, "B"): "Engine!B4"}
