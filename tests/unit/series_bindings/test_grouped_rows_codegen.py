"""Setter codegen round-trip tests for grouped-row matrix bindings."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import cast

import pytest

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)
from excel_grapher.series_bindings.setter_codegen import (
    emit_input_coerce_helpers,
    emit_setter_function,
    emit_setter_helpers,
    emit_setters_block,
)
from tests.fixtures.series_bindings.grouped_matrix_helpers import (
    MATRIX_GROUPED_ROWS_BINDINGS,
    write_grouped_matrix_workbook,
)


def _exec_setters(lines: list[str]) -> dict[str, object]:
    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
    }
    source_lines = lines
    if "def coerce_setter_input(" not in "\n".join(lines):
        source_lines = emit_input_coerce_helpers() + emit_setter_helpers() + lines
    exec("\n".join(source_lines), namespace)
    return namespace


def test_grouped_rows_matrix_setter_round_trip(tmp_path: Path) -> None:
    wb_path = tmp_path / "grouped_inputs.xlsx"
    write_grouped_matrix_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!C2:D8"),
        load_values=True,
    )
    bindings = load_series_bindings(MATRIX_GROUPED_ROWS_BINDINGS)
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["ok"] is True, resolved["issues"]

    with pytest.warns(UserWarning, match="read_discrete_risks_range"):
        code = "\n".join(emit_setters_block(graph, wb_path, bindings))
    assert code.count("def set_discrete_risks(") == 1
    assert "def read_discrete_risks_range(" not in code

    ns = _exec_setters(emit_setter_function(series, resolved))
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_discrete_risks"],
    )
    ctx = EvalContext(
        inputs=coerce_inputs_dict({"Inputs!C3": 1.1}),
        resolver=lambda _a: None,
    )
    setter(
        ctx,
        [
            {
                "SCENARIO": "Paris",
                "SHOCK_TYPE": "Primary expenditure",
                "TIME_PERIOD": 2025,
                "OBS_VALUE": 9.9,
            },
            {
                "SCENARIO": "Moderate",
                "SHOCK_TYPE": "Revenue",
                "TIME_PERIOD": 2024,
                "OBS_VALUE": 8.8,
            },
        ],
    )
    assert ctx.inputs["Inputs!D4"] == 9.9
    assert ctx.inputs["Inputs!C7"] == 8.8
    assert ctx.inputs["Inputs!C3"] == 1.1
