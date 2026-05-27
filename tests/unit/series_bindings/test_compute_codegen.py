"""Unit tests for generated series-binding output compute functions."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict, xl_cell
from excel_grapher.series_bindings import (
    Records,
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)
from excel_grapher.series_bindings.compute_codegen import emit_compute_function

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def _write_formula_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 5.0)
    ws.write_formula("C2", "=B2*2")
    wb.close()


def _exec_compute(
    lines: list[str],
    *,
    resolver: Callable[[str], Any],
) -> dict[str, object]:
    import warnings

    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
        "xl_cell": xl_cell,
        "warnings": warnings,
        "make_context": lambda inputs=None: EvalContext(
            inputs=coerce_inputs_dict(inputs or {}),
            resolver=resolver,
        ),
    }
    exec("\n".join(lines), namespace)
    return namespace


def test_emit_compute_returns_records_with_obs_value(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula.xlsx"
    _write_formula_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2"], load_values=True)

    series = {
        "id": "scaled_output",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_scaled_output"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
            "dimensions": [
                {
                    "concept": "LABEL",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "scaled"},
                }
            ],
        },
        "key": ["LABEL"],
    }
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    lines = [
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
        *emit_compute_function(series, resolved),
    ]

    formula_impl = graph.get_node("Sheet1!C2")
    assert formula_impl is not None

    def resolver(address: str):
        if address == "Sheet1!C2":
            return lambda ctx: 10.0
        return None

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[..., Records], ns["compute_scaled_output"])
    records = compute()
    assert len(records) == 1
    assert records[0]["LABEL"] == "scaled"
    assert records[0]["OBS_VALUE"] == 10.0


def test_emit_compute_borvelia_includes_frequency(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    from tests.unit.series_bindings.test_resolve import _write_borvelia_workbook

    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "shard_borvelia_output.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    lines = [
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
        *emit_compute_function(series, resolved),
    ]

    def resolver(address: str):
        return lambda ctx: xl_cell(ctx, address)

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[..., Records], ns["compute_borvelia_primary_balance"])
    records = compute()
    assert len(records) == 5
    assert all(r["FREQUENCY"] == "A" for r in records)
    assert {r["TIME_PERIOD"] for r in records} == {1, 2, 3, 4, 5}
