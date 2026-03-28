from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any, cast

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.codegen import CodeGenerator
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.export_runtime.cache import CircularReferenceWarning
from excel_grapher.evaluator.name_utils import parse_address
from tests.evaluator.parity_harness import exec_generated_code

if TYPE_CHECKING:
    import numpy as np

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


def test_parity_direct_self_cycle_returns_zero() -> None:
    graph = _make_graph(_make_node("S!A1", "=S!A1", None))

    with FormulaEvaluator(graph) as ev, pytest.warns(CircularReferenceWarning):
        evaluator_result = ev.evaluate(["S!A1"])["S!A1"]

    with pytest.warns(RuntimeWarning, match=r"Circular reference detected; returning 0") as w:
        generated_results, _code, _ns = exec_generated_code(graph, ["S!A1"])
        generated_result = generated_results["S!A1"]
    assert any(wi.category.__name__ == "CircularReferenceWarning" for wi in w)

    assert evaluator_result == 0
    assert generated_result == 0


def test_parity_indirect_cycle_returns_zero() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=S!B1", None),
        _make_node("S!B1", "=S!A1", None),
    )

    with FormulaEvaluator(graph) as ev, pytest.warns(CircularReferenceWarning):
        evaluator_result = ev.evaluate(["S!A1", "S!B1"])

    with pytest.warns(RuntimeWarning, match=r"Circular reference detected; returning 0") as w:
        generated_result, _code, _ns = exec_generated_code(graph, ["S!A1", "S!B1"])
    assert any(wi.category.__name__ == "CircularReferenceWarning" for wi in w)

    assert evaluator_result == {"S!A1": 0, "S!B1": 0}
    if "S!A1:S!B1" in generated_result:
        result = cast("np.ndarray", generated_result["S!A1:S!B1"])
        assert result.tolist() == [[0, 0]]
    else:
        assert generated_result == {"S!A1": 0, "S!B1": 0}


def test_iterative_self_cycle_converges_with_parity() -> None:
    graph = _make_graph(_make_node("S!A1", "=(S!A1+1)/2", None))
    targets = ["S!A1"]

    with FormulaEvaluator(
        graph, iterate_enabled=True, iterate_count=100, iterate_delta=1e-6
    ) as ev:
        evaluator_result = ev.evaluate(targets)

    generated_code = CodeGenerator(
        graph, iterate_enabled=True, iterate_count=100, iterate_delta=1e-6
    ).generate(targets)
    ns: dict[str, Any] = {}
    exec(generated_code, ns)
    compute_all = cast(Callable[[], dict[str, float]], ns["compute_all"])
    generated_result = compute_all()

    assert abs(float(cast(Any, evaluator_result["S!A1"])) - 1.0) <= 1e-4
    assert abs(float(generated_result["S!A1"]) - 1.0) <= 1e-4
    assert abs(float(cast(Any, evaluator_result["S!A1"])) - float(generated_result["S!A1"])) <= 1e-9


def test_iterative_mutual_cycle_converges_with_parity() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=(S!B1+1)/2", None),
        _make_node("S!B1", "=(S!A1+1)/2", None),
    )
    targets = ["S!A1", "S!B1"]

    with FormulaEvaluator(
        graph, iterate_enabled=True, iterate_count=100, iterate_delta=1e-6
    ) as ev:
        evaluator_result = ev.evaluate(targets)

    generated_code = CodeGenerator(
        graph, iterate_enabled=True, iterate_count=100, iterate_delta=1e-6
    ).generate(targets)
    ns: dict[str, Any] = {}
    exec(generated_code, ns)
    compute_all = cast(Callable[[], dict[str, Any]], ns["compute_all"])
    generated_raw = compute_all()
    if "S!A1:S!B1" in generated_raw:
        result = cast("np.ndarray", generated_raw["S!A1:S!B1"])
        generated_result = {"S!A1": result.tolist()[0][0], "S!B1": result.tolist()[0][1]}
    else:
        generated_result = cast("dict[str, float]", generated_raw)

    for key in targets:
        assert abs(float(cast(Any, evaluator_result[key])) - 1.0) <= 1e-4
        assert abs(float(generated_result[key]) - 1.0) <= 1e-4
        assert abs(float(cast(Any, evaluator_result[key])) - float(generated_result[key])) <= 1e-6


def test_iterative_max_iterations_respected_for_oscillation() -> None:
    graph = _make_graph(_make_node("S!A1", "=1-S!A1", None))
    targets = ["S!A1"]

    with FormulaEvaluator(
        graph, iterate_enabled=True, iterate_count=3, iterate_delta=1e-12
    ) as ev:
        evaluator_result = ev.evaluate(targets)

    generated_code = CodeGenerator(
        graph, iterate_enabled=True, iterate_count=3, iterate_delta=1e-12
    ).generate(targets)
    ns: dict[str, Any] = {}
    exec(generated_code, ns)
    compute_all = cast(Callable[[], dict[str, float]], ns["compute_all"])
    generated_result = compute_all()

    assert evaluator_result["S!A1"] == generated_result["S!A1"]


def test_iterative_lazy_if_avoids_cycle_when_branch_not_taken() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=IF(S!C1=0,S!B1,5)", None),
        _make_node("S!B1", "=S!A1", None),
        _make_node("S!C1", None, 1),
    )
    targets = ["S!A1"]

    with FormulaEvaluator(
        graph, iterate_enabled=True, iterate_count=10, iterate_delta=1e-9
    ) as ev:
        evaluator_result = ev.evaluate(targets)

    generated_code = CodeGenerator(
        graph, iterate_enabled=True, iterate_count=10, iterate_delta=1e-9
    ).generate(targets)
    ns: dict[str, Any] = {}
    exec(generated_code, ns)
    compute_all = cast(Callable[[], dict[str, float]], ns["compute_all"])
    generated_result = compute_all()

    assert evaluator_result["S!A1"] == 5
    assert generated_result["S!A1"] == 5

