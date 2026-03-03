"""Parity tests for INDEX edge cases."""

from excel_grapher import DependencyGraph
from excel_grapher import Node

from excel_grapher.evaluator.name_utils import parse_address
from excel_grapher.evaluator.types import XlError
from tests.evaluator.parity_harness import assert_codegen_matches_evaluator


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


def test_index_parity_with_non_array_input() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=INDEX(TRUE,1)", None),
        _make_node("S!A2", "=INDEX(1,1)", None),
        _make_node("S!A3", '=INDEX("text",1)', None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2", "S!A3"])
    assert result.generated_results["S!A1"] == XlError.VALUE
    assert result.generated_results["S!A2"] == XlError.VALUE
    assert result.generated_results["S!A3"] == XlError.VALUE
