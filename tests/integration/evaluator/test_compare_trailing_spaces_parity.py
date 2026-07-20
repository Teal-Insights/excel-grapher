"""Trailing-space text compare: evaluator ↔ export parity (GitHub #434)."""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.types import XlError
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


def test_compare_trailing_spaces_codegen_parity() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, "High"),
        _make_node("S!A2", None, "High "),
        _make_node("S!B1", '="High"="High "', None),
        _make_node("S!B2", '="High"<>"High "', None),
        _make_node("S!B3", "=S!A1=S!A2", None),
        _make_node("S!B4", '="high"="HIGH"', None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2", "S!B3", "S!B4"])
    assert result.generated_results["S!B1"] is False
    assert result.generated_results["S!B2"] is True
    assert result.generated_results["S!B3"] is False
    assert result.generated_results["S!B4"] is True


def test_match_exact_trailing_spaces_codegen_parity() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, "High"),
        _make_node("S!A2", None, "High "),
        _make_node("S!B1", '=MATCH("High",S!A1:S!A1,0)', None),
        _make_node("S!B2", '=MATCH("High",S!A2:S!A2,0)', None),
        _make_node("S!B3", '=MATCH("High ",S!A2:S!A2,0)', None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2", "S!B3"])
    assert result.generated_results["S!B1"] == 1
    assert result.generated_results["S!B2"] == XlError.NA
    assert result.generated_results["S!B3"] == 1
