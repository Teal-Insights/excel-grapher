"""Parity tests for missing cell references."""

import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.evaluator.name_utils import parse_address
from tests.evaluator.parity_harness import exec_generated_code


def _make_node(address: str, formula: str | None, value: object) -> Node:
    """Helper to create a Node from a sheet-qualified address."""
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
    """Helper to create a DependencyGraph from nodes."""
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_missing_cell_reference_raises_in_both_paths() -> None:
    """Missing cell access should raise in evaluator and generated code."""
    graph = _make_graph(
        _make_node("S!A1", "=S!B1+1", None),
        # S!B1 is NOT in the graph
    )

    with pytest.raises(KeyError, match="S!B1"), FormulaEvaluator(graph) as ev:
        ev.evaluate(["S!A1"])

    with pytest.raises(KeyError, match="S!B1"):
        exec_generated_code(graph, ["S!A1"])
