"""Parity tests for MATCH."""

from excel_grapher import DependencyGraph, Node
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


def test_match_parity_with_single_cell() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=MATCH(1, S!A1, 0)", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == 1


def test_match_parity_with_error_single_cell() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, XlError.VALUE),
        _make_node("S!B1", "=MATCH(TRUE, S!A1, 0)", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == XlError.VALUE
