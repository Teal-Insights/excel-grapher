"""Circular references: FormulaEvaluator cycle and iteration behavior (integration)."""

from __future__ import annotations

from typing import Any, cast

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.runtime.cache import CircularReferenceWarning


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


def _make_graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_direct_self_cycle_returns_zero() -> None:
    graph = _make_graph(_make_node("S!A1", "=S!A1", None))
    with FormulaEvaluator(graph) as ev, pytest.warns(CircularReferenceWarning):
        evaluator_result = ev.evaluate(["S!A1"])["S!A1"]
    assert evaluator_result == 0


def test_indirect_cycle_returns_zero() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=S!B1", None),
        _make_node("S!B1", "=S!A1", None),
    )
    with FormulaEvaluator(graph) as ev, pytest.warns(CircularReferenceWarning):
        evaluator_result = ev.evaluate(["S!A1", "S!B1"])
    assert evaluator_result == {"S!A1": 0, "S!B1": 0}


def test_iterative_self_cycle_converges() -> None:
    graph = _make_graph(_make_node("S!A1", "=(S!A1+1)/2", None))
    with FormulaEvaluator(graph, iterate_enabled=True, iterate_count=100, iterate_delta=1e-6) as ev:
        evaluator_result = ev.evaluate(["S!A1"])
    assert abs(float(cast(Any, evaluator_result["S!A1"])) - 1.0) <= 1e-4


def test_iterative_mutual_cycle_converges() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=(S!B1+1)/2", None),
        _make_node("S!B1", "=(S!A1+1)/2", None),
    )
    with FormulaEvaluator(graph, iterate_enabled=True, iterate_count=100, iterate_delta=1e-6) as ev:
        evaluator_result = ev.evaluate(["S!A1", "S!B1"])
    for key in ("S!A1", "S!B1"):
        assert abs(float(cast(Any, evaluator_result[key])) - 1.0) <= 1e-4


def test_iterative_max_iterations_respected_for_oscillation() -> None:
    graph = _make_graph(_make_node("S!A1", "=1-S!A1", None))
    with FormulaEvaluator(graph, iterate_enabled=True, iterate_count=3, iterate_delta=1e-12) as ev:
        evaluator_result = ev.evaluate(["S!A1"])
    assert evaluator_result["S!A1"] in {0, 1}


def test_iterative_lazy_if_avoids_cycle_when_branch_not_taken() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=IF(S!C1=0,S!B1,5)", None),
        _make_node("S!B1", "=S!A1", None),
        _make_node("S!C1", None, 1),
    )
    with FormulaEvaluator(graph, iterate_enabled=True, iterate_count=10, iterate_delta=1e-9) as ev:
        evaluator_result = ev.evaluate(["S!A1"])
    assert evaluator_result["S!A1"] == 5
