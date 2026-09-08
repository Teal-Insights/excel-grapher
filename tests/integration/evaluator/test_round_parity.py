"""ROUND: evaluator and generated export runtime agree on synthetic graphs (integration).

Live Excel parity for ``ROUND`` lives in ``test_round_excel_parity.py`` (slow, run-if-available).
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


def test_round_parity_excel_half_away_from_zero() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1.25),
        _make_node("S!B1", "=ROUND(1.25,1)", None),
        _make_node("S!B2", "=ROUND(-1.25,1)", None),
        _make_node("S!B3", "=ROUND(2.5,0)", None),
        _make_node("S!B4", "=ROUND(125,-1)", None),
        _make_node("S!B5", "=ROUND(S!A1,1)", None),
    )

    results = evaluate_targets(graph, ["S!B1", "S!B2", "S!B3", "S!B4", "S!B5"])
    assert results["S!B1"] == 1.3
    assert results["S!B2"] == -1.3
    assert results["S!B3"] == 3.0
    assert results["S!B4"] == 130.0
    assert results["S!B5"] == 1.3
