"""EXP: evaluator and generated export runtime agree on synthetic graphs (integration)."""

from __future__ import annotations

import math

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


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


def test_exp_parity_scalar_and_cell_ref() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=EXP(1)", None),
        _make_node("S!B2", "=EXP(S!A1)", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2"])
    assert result.generated_results["S!B1"] == pytest.approx(math.e)
    assert result.generated_results["S!B2"] == pytest.approx(math.e)
    assert "xl_exp" in result.generated_code


def test_exp_parity_logistic_convergence_pattern() -> None:
    """Mirror Q-CRAFT logistic convergence with a computed year column (issue #333 MCVE)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=EXP(S!A1)", None),
        _make_node("S!A2", None, 0.5),
        _make_node("S!A3", None, 15.0),
        _make_node("S!B2", "=1/(1+EXP(-S!A2*(S!B1-S!A3)))", None),
    )

    expected = 1.0 / (1.0 + math.exp(-0.5 * (math.e - 15.0)))
    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2"])
    assert result.generated_results["S!B1"] == pytest.approx(math.e)
    assert result.generated_results["S!B2"] == pytest.approx(expected)
