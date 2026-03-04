"""Parity tests for NUMBERVALUE."""

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


def test_numbervalue_parity_with_explicit_separators() -> None:
    graph = _make_graph(
        _make_node(
            "S!A1",
            '=_xlfn.NUMBERVALUE("1,234.56", ".", ",")',
            None,
        ),
        _make_node(
            "S!A2",
            '=NUMBERVALUE("2.500,00", ",", ".")',
            None,
        ),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2"])
    assert result.generated_results["S!A1"] == 1234.56
    assert result.generated_results["S!A2"] == 2500.0


def test_numbervalue_parity_with_percent_and_parentheses() -> None:
    graph = _make_graph(
        _make_node(
            "S!B1",
            '=_xlfn.NUMBERVALUE("12%", ".", ",")',
            None,
        ),
        _make_node(
            "S!B2",
            '=_xlfn.NUMBERVALUE("(1,234.56)", ".", ",")',
            None,
        ),
        _make_node(
            "S!B3",
            '=NUMBERVALUE("12,5%", ",", ".")',
            None,
        ),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2", "S!B3"])
    assert result.generated_results["S!B1"] == 0.12
    assert result.generated_results["S!B2"] == -1234.56
    assert result.generated_results["S!B3"] == 0.125


def test_numbervalue_parity_with_currency_and_spaces() -> None:
    graph = _make_graph(
        _make_node(
            "S!C1",
            '=_xlfn.NUMBERVALUE("$1,234.56", ".", ",")',
            None,
        ),
        _make_node(
            "S!C2",
            '=_xlfn.NUMBERVALUE(" -$1,234.56 ", ".", ",")',
            None,
        ),
        _make_node(
            "S!C3",
            '=_xlfn.NUMBERVALUE("€1 234,56", ",", " ")',
            None,
        ),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!C1", "S!C2", "S!C3"])
    assert result.generated_results["S!C1"] == 1234.56
    assert result.generated_results["S!C2"] == -1234.56
    assert result.generated_results["S!C3"] == 1234.56
