"""Tests for keyless scalar series bindings (issue #215)."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict
from excel_grapher.series_bindings import (
    expand_data_range,
    resolve_series_binding,
    validate_bindings_document,
)
from excel_grapher.series_bindings.setter_codegen import emit_setter_function

KEYLESS_SCALAR_BINDING: dict[str, Any] = {
    "schema_version": "1.2.0",
    "series": [
        {
            "id": "country_name",
            "sheet": "Inputs",
            "data_range": "Inputs!B5",
            "layout": "scalar",
            "input": {
                "setter": {
                    "name": "set_country_name",
                    "record_contract": "records",
                    "strict": True,
                }
            },
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "string",
                    "bind": {"kind": "data_cell", "read": "string"},
                },
                "dimensions": [],
                "attributes": [
                    {
                        "concept": "PARAMETER",
                        "role": "attribute",
                        "value": "country_name",
                        "include_in_record": False,
                    }
                ],
            },
            "key": [],
            "series_context": {"PARAMETER": "country_name"},
        }
    ],
}


def _exec_setters(lines: list[str]) -> dict[str, object]:
    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
    }
    exec("\n".join(lines), namespace)
    return namespace


def test_keyless_scalar_binding_passes_schema() -> None:
    bindings = validate_bindings_document(KEYLESS_SCALAR_BINDING)
    series = bindings["series"][0]
    assert series["key"] == []
    assert series["structure"]["dimensions"] == []


def test_keyless_scalar_resolves_single_leaf(tmp_path: Path) -> None:
    wb_path = tmp_path / "inputs.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write("B5", "Litellia")
    wb.close()

    graph = create_dependency_graph(wb_path, ["Inputs!B5"], load_values=True)
    series = cast(dict[str, Any], KEYLESS_SCALAR_BINDING["series"][0])
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True
    assert resolved["requires_address"] is False
    assert len(resolved["leaves"]) == 1
    assert resolved["leaves"][0]["key"] == {}
    assert resolved["leaves"][0]["record"]["OBS_VALUE"] == "Litellia"


def test_keyless_scalar_multi_leaf_fails_resolution(tmp_path: Path) -> None:
    wb_path = tmp_path / "multi.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write("B5", "A")
    ws.write("C5", "B")
    wb.close()

    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!B5:C5"),
        load_values=True,
    )
    series: dict[str, Any] = {
        **cast(dict[str, Any], KEYLESS_SCALAR_BINDING["series"][0]),
        "data_range": "Inputs!B5:C5",
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is False
    codes = {issue["code"] for issue in resolved["issues"]}
    assert "keyless_scalar_ambiguous" in codes


def test_keyless_scalar_setter_accepts_shorthand_value(tmp_path: Path) -> None:
    wb_path = tmp_path / "inputs.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write("B5", "Litellia")
    wb.close()

    graph = create_dependency_graph(wb_path, ["Inputs!B5"], load_values=True)
    series = cast(dict[str, Any], KEYLESS_SCALAR_BINDING["series"][0])
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, object], None],
        ns["set_country_name"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, "Newland")
    assert ctx.inputs["Inputs!B5"] == "Newland"


def test_keyless_scalar_setter_accepts_record_without_keys(tmp_path: Path) -> None:
    wb_path = tmp_path / "inputs.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write("B5", "Litellia")
    wb.close()

    graph = create_dependency_graph(wb_path, ["Inputs!B5"], load_values=True)
    series = cast(dict[str, Any], KEYLESS_SCALAR_BINDING["series"][0])
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, object], None],
        ns["set_country_name"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"OBS_VALUE": "Newland"}])
    assert ctx.inputs["Inputs!B5"] == "Newland"
