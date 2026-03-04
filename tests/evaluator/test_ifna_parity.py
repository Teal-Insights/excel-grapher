"""Parity tests for IFNA."""

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.name_utils import parse_address
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


def test_ifna_parity_with_na_and_value() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=NA()", None),
        _make_node("S!A2", "=IFNA(S!A1, 7)", None),
        _make_node("S!A3", "=_xlfn.IFNA(S!A1, 9)", None),
        _make_node("S!A4", "=IFNA(5, 11)", None),
        _make_node("S!A5", '=_xlfn.IFNA("text", "fallback")', None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!A2", "S!A3", "S!A4", "S!A5"])
    assert result.generated_results["S!A2"] == 7
    assert result.generated_results["S!A3"] == 9
    assert result.generated_results["S!A4"] == 5
    assert result.generated_results["S!A5"] == "text"
