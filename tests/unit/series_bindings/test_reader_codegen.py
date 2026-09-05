"""Unit tests for generated series-binding read_* duals of set_* setters."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict, xl_cell
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)
from excel_grapher.series_bindings.setter_codegen import (
    emit_input_coerce_helpers,
    emit_reader_function,
    emit_setter_function,
    emit_setter_helpers,
    emit_setters_block,
)
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


def _exec_readers(
    lines: list[str],
    *,
    extra: dict[str, object] | None = None,
) -> dict[str, object]:
    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
        "xl_cell": xl_cell,
        "CellValue": object,
    }
    if extra:
        namespace.update(extra)
    source_lines = lines
    if "def coerce_setter_input(" not in "\n".join(lines):
        source_lines = emit_input_coerce_helpers() + emit_setter_helpers() + lines
    exec("\n".join(source_lines), namespace)
    return namespace


def test_emit_reader_returns_value_by_key(tmp_path: Path) -> None:
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
    lines = emit_setter_function(series, resolved) + emit_reader_function(series, resolved)
    ns = _exec_readers(lines)
    reader = cast(Callable[..., object], ns["read_borvelia_primary_balance"])

    ctx = EvalContext(
        inputs=coerce_inputs_dict({"Inputs!H5": 42.0}),
        resolver=lambda _a: None,
    )
    assert reader(ctx, time_period=3) == 42.0


def test_emit_reader_missing_key_raises(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    ns = _exec_readers(
        emit_setter_function(series, resolved) + emit_reader_function(series, resolved)
    )
    reader = cast(Callable[..., object], ns["read_borvelia_primary_balance"])

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    with pytest.raises(ValueError, match="no leaf matches key"):
        reader(ctx, time_period=99)


def test_emit_reader_uses_snake_case_key_params(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_reader_function(series, resolved)
    source = "\n".join(lines)
    assert "def read_borvelia_primary_balance(" in source
    assert "*," in source
    assert "time_period:" in source
    assert "TIME_PERIOD" not in source.split("def read_borvelia_primary_balance(")[1].split(")")[0]


def test_emit_reader_keyless_scalar(tmp_path: Path) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write("B5", "Litellia")
    wb.close()

    graph = create_dependency_graph(wb_path, ["Inputs!B5"], load_values=True)
    series = {
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
        },
        "key": [],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved) + emit_reader_function(series, resolved)
    ns = _exec_readers(lines)
    reader = cast(Callable[..., object], ns["read_country_name"])

    ctx = EvalContext(
        inputs=coerce_inputs_dict({"Inputs!B5": "Litellia"}),
        resolver=lambda _a: None,
    )
    assert reader(ctx) == "Litellia"
    source = "\n".join(lines)
    assert "def read_country_name(ctx: EvalContext) -> CellValue:" in source or (
        "def read_country_name(" in source and "time_period" not in source
    )


def test_emit_reader_requires_address_uses_address_kwarg(tmp_path: Path) -> None:
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
    lines = emit_setter_function(series, resolved) + emit_reader_function(series, resolved)
    source = "\n".join(lines)
    assert "def read_dup_headers(ctx: EvalContext, *, address: str) -> CellValue:" in source
    assert "_READER_ADDRESSES_DUP_HEADERS" in source

    ns = _exec_readers(lines)
    reader = cast(Callable[..., object], ns["read_dup_headers"])
    ctx = EvalContext(
        inputs=coerce_inputs_dict({"Sheet1!C2": 10.0, "Sheet1!D2": 20.0}),
        resolver=lambda _a: None,
    )
    assert reader(ctx, address="Sheet1!C2") == 10.0
    assert reader(ctx, address="Sheet1!D2") == 20.0
    with pytest.raises(ValueError, match="not a leaf"):
        reader(ctx, address="Sheet1!Z9")


def test_emit_setters_block_includes_readers(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    lines = emit_setters_block(graph, wb_path, bindings)
    source = "\n".join(lines)
    assert "def set_borvelia_primary_balance(" in source
    assert "def read_borvelia_primary_balance(" in source
    assert "def read_borvelia_primary_balance_range(" in source


def test_emit_reader_range_returns_values(tmp_path: Path) -> None:
    from excel_grapher.exporter.export_runtime.offset import xl_range

    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    from excel_grapher.series_bindings.setter_codegen import emit_reader_range_function

    lines = (
        emit_setter_function(series, resolved)
        + emit_reader_function(series, resolved)
        + emit_reader_range_function(series, resolved)
    )
    ns = _exec_readers(lines, extra={"xl_range": xl_range})
    reader_range = cast(Callable[..., object], ns["read_borvelia_primary_balance_range"])

    values = {
        "Inputs!F5": -2.0,
        "Inputs!G5": -1.0,
        "Inputs!H5": 0.0,
        "Inputs!I5": 1.0,
        "Inputs!J5": 2.0,
    }
    ctx = EvalContext(inputs=coerce_inputs_dict(values), resolver=lambda _a: None)
    result = cast(Any, reader_range(ctx))
    assert result.cell(1, 1) == -2.0
    assert result.cell(1, 5) == 2.0
