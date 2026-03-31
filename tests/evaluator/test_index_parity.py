"""Parity tests for INDEX edge cases."""

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


def test_index_omit_row_returns_column_for_match() -> None:
    """INDEX(range,,k) returns column k; used by LIC-DSF classification table."""
    graph = _make_graph(
        _make_node("S!A1", "1", None),
        _make_node("S!A2", "4", None),
        _make_node("S!A3", "7", None),
        _make_node("S!B1", "2", None),
        _make_node("S!B2", "5", None),
        _make_node("S!B3", "8", None),
        _make_node("S!C1", "3", None),
        _make_node("S!C2", "6", None),
        _make_node("S!C3", "9", None),
        _make_node("S!D1", "5", None),
        _make_node("S!E1", "=MATCH(S!D1, INDEX(S!A1:S!C3,,2), 0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!E1"])
    assert result.evaluator_results["S!E1"] == 2
    assert result.generated_results["S!E1"] == 2


def test_index_omit_col_returns_row_for_match() -> None:
    """INDEX(range,k,) returns row k for 2-D arrays."""
    graph = _make_graph(
        _make_node("S!A1", "1", None),
        _make_node("S!A2", "4", None),
        _make_node("S!A3", "7", None),
        _make_node("S!B1", "2", None),
        _make_node("S!B2", "5", None),
        _make_node("S!B3", "8", None),
        _make_node("S!C1", "3", None),
        _make_node("S!C2", "6", None),
        _make_node("S!C3", "9", None),
        _make_node("S!D1", "5", None),
        _make_node("S!E1", "=MATCH(S!D1, INDEX(S!A1:S!C3,2,), 0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!E1"])
    assert result.evaluator_results["S!E1"] == 2
    assert result.generated_results["S!E1"] == 2
