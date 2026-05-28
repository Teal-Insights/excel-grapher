"""Unit tests for generated series-binding setters."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import cast

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)
from excel_grapher.series_bindings.docstrings import (
    FieldDoc,
    SeriesFunctionDoc,
    register_series_docstring_callback,
)
from excel_grapher.series_bindings.setter_codegen import emit_setter_function, emit_setters_block

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


def _exec_setters(
    lines: list[str],
    *,
    extra: dict[str, object] | None = None,
) -> dict[str, object]:
    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
    }
    if extra:
        namespace.update(extra)
    exec("\n".join(lines), namespace)
    return namespace


def test_emit_setter_updates_context_by_key(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"TIME_PERIOD": 3, "OBS_VALUE": 42.0}])
    assert ctx.inputs["Inputs!H5"] == 42.0


def test_emit_setter_missing_key_raises(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    ns = _exec_setters(emit_setter_function(series, resolved))
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    with pytest.raises(ValueError, match="missing key fields"):
        setter(ctx, [{"OBS_VALUE": 1.0}])


def test_emit_setter_missing_measure_raises(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    ns = _exec_setters(emit_setter_function(series, resolved))
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    with pytest.raises(ValueError, match="missing required field 'OBS_VALUE'"):
        setter(ctx, [{"TIME_PERIOD": 3}])


def test_emit_setter_allow_address(tmp_path: Path) -> None:
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
        "setter": {"name": "set_dup_headers", "allow_address": True, "strict": False},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
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
    lines = emit_setter_function(series, resolved)
    ns = _exec_setters(lines)
    setter = cast(Callable[[EvalContext, list[dict[str, object]]], None], ns["set_dup_headers"])
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"address": "Sheet1!D2", "OBS_VALUE": 99.0}])
    assert ctx.inputs["Sheet1!D2"] == 99.0


def test_emit_setters_block_all_series(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    lines = emit_setters_block(graph, wb_path, bindings)
    assert "def set_borvelia_primary_balance(" in "\n".join(lines)


def test_emit_setters_block_skips_series_without_graph_leaf_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!A2"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    with pytest.warns(UserWarning, match="No resolved input cells"):
        lines = emit_setters_block(graph, wb_path, bindings)

    assert "def set_borvelia_primary_balance(" not in "\n".join(lines)


def test_emit_setter_structured_docstring_callback(tmp_path: Path) -> None:
    callback_name = "_test_setter_structured_docstring"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary="Set borvelia values.",
            purpose="Updates borvelia primary balance inputs.",
            record_matching="Match records by TIME_PERIOD.",
            field_descriptions={
                "TIME_PERIOD": FieldDoc(description="Reporting year."),
                "OBS_VALUE": FieldDoc(description='Value with "quotes".'),
            },
        ),
    )
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(
        series,
        resolved,
        graph=graph,
        workbook=wb_path,
        bindings=bindings,
        series_docstring_callback=callback_name,
    )
    code = "\n".join(lines)
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    assert setter.__doc__ is not None
    assert "Set borvelia values." in setter.__doc__
    assert 'Value with "quotes".' in setter.__doc__
    exec(code, {"EvalContext": EvalContext, "coerce_inputs_dict": coerce_inputs_dict})


def test_emit_setter_docstring_callback_none_omits_docstring(tmp_path: Path) -> None:
    callback_name = "_test_setter_none_docstring"
    register_series_docstring_callback(callback_name, lambda ctx: None)
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(
        series,
        resolved,
        graph=graph,
        workbook=wb_path,
        bindings=bindings,
        series_docstring_callback=callback_name,
    )
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    assert setter.__doc__ is None


def test_emit_setter_callback_requires_codegen_context(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    with pytest.raises(ValueError, match="requires graph, workbook, and bindings"):
        _ = emit_setter_function(
            series,
            resolved,
            series_docstring_callback="series_notes",
        )
