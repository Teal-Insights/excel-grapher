"""VALUE: evaluator and generated export runtime agree on synthetic graphs (integration).

Guards numeric parsing parity through `assert_codegen_matches_evaluator` for small
dependency graphs.
"""

from excel_grapher import DependencyGraph, Node, XlError
from excel_grapher.core.address_keys import parse_address
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


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


def test_value_parity_with_lic_dsf_style_key() -> None:
    graph = _make_graph(
        _make_node("S!A1", '=VALUE("6522014")', None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!A1"])
    assert result.generated_results["S!A1"] == 6522014.0


def test_value_parity_with_currency() -> None:
    graph = _make_graph(
        _make_node("S!B1", '=VALUE("$1,234.56")', None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == 1234.56


def test_value_parity_with_invalid_text() -> None:
    graph = _make_graph(
        _make_node("S!C1", '=VALUE("abc")', None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!C1"])
    assert result.generated_results["S!C1"] == XlError.VALUE
