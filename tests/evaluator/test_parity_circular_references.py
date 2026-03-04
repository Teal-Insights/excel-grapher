from __future__ import annotations

from typing import TYPE_CHECKING, cast

import pytest

from excel_grapher import DependencyGraph, Node
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

