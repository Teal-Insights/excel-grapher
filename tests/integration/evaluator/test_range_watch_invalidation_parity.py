"""Evaluator ↔ export parity for range-watch invalidation after `set_inputs` (#583)."""

from __future__ import annotations

from collections.abc import Callable

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.core import CellValue
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.export_runtime.math import xl_sum
from excel_grapher.exporter.export_runtime.offset import xl_range
from excel_grapher.runtime.cache import EvalContext, xl_cell


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _sum_range_graph() -> DependencyGraph:
    graph = DependencyGraph()
    for row in range(1, 6):
        graph.add_node(_make_node(f"S!A{row}", None, float(row)))
    graph.add_node(_make_node("S!B1", "=SUM(S!A1:S!A5)", 15.0))
    return graph


def _evaluator_sequence() -> list[object]:
    graph = _sum_range_graph()
    values: list[object] = []
    with FormulaEvaluator(graph) as ev:
        values.append(ev.evaluate("S!B1"))
        graph.set_node_value("S!A3", 30.0)
        values.append(ev.evaluate("S!B1"))
        graph.set_node_value("S!A3", 3.0)
        values.append(ev.evaluate("S!B1"))
    return values


def _exported_runtime_resolver() -> Callable[[str], Callable[[EvalContext], CellValue] | None]:
    def _b1(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_range(ctx, "S!A1:A5"))

    return {"S!B1": _b1}.get


def _exported_sequence() -> list[object]:
    ctx = EvalContext(
        inputs={f"S!A{row}": float(row) for row in range(1, 6)},
        resolver=_exported_runtime_resolver(),
    )
    values: list[object] = []
    values.append(xl_cell(ctx, "S!B1"))
    ctx.set_inputs({"S!A3": 30.0})
    values.append(xl_cell(ctx, "S!B1"))
    ctx.set_inputs({"S!A3": 3.0})
    values.append(xl_cell(ctx, "S!B1"))
    return values


def test_sum_range_invalidation_matches_evaluator() -> None:
    assert _evaluator_sequence() == _exported_sequence() == [15.0, 42.0, 15.0]
