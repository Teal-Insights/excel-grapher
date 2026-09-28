"""CHOOSE: evaluator and generated export runtime agree on synthetic graphs (integration).

Excel truncates a fractional index (`CHOOSE(1.5, a, b)` is `a`).
"""

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from tests.integration.utils.parity_harness import evaluate_targets


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


def test_choose_parity_truncates_fractional_index() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1.5),
        _make_node("S!A2", None, 2.99),
        _make_node("S!B1", "=CHOOSE(S!A1,10,20,30)", None),
        _make_node("S!B2", "=CHOOSE(S!A2,10,20,30)", None),
        _make_node("S!B3", "=CHOOSE(1.5,10,20,30)", None),
    )

    results = evaluate_targets(graph, ["S!B1", "S!B2", "S!B3"])
    assert results["S!B1"] == 10
    assert results["S!B2"] == 20
    assert results["S!B3"] == 10
