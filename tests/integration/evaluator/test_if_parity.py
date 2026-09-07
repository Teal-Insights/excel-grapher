"""Array-style SUM/AVERAGE/MAX(IF) and SUMPRODUCT(IF): evaluator and codegen agree.

Live Excel parity lives in `test_if_excel_parity.py` (slow, run-if-available).
"""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
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


def test_sum_if_array_codegen_parity() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, -1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!B3", None, 30),
        _make_node("S!C1", None, 100),
        _make_node("S!C2", None, 200),
        _make_node("S!C3", None, 300),
        _make_node("S!E1", None, 2),
        _make_node("S!D1", "=SUM(IF(S!A1:A3>0,S!A1:A3))", None),
        _make_node("S!D2", "=SUM(IF(S!A1:A3>0,S!B1:B3,0))", None),
        _make_node("S!D3", "=SUM(IF(S!A1:A3>0,S!B1:B3,S!C1:C3))", None),
        _make_node("S!D4", "=SUM(IF(S!A1:A3=S!E1,S!B1:B3,0))", None),
        _make_node("S!D5", "=SUM(IF(S!A1:A3>0,IF(S!B1:B3>15,S!C1:C3,0),0))", None),
        _make_node("S!D6", "=SUMPRODUCT(IF(S!A1:A3>0,S!B1:B3,0))", None),
        _make_node("S!D7", "=AVERAGE(IF(S!A1:A3>0,S!A1:A3))", None),
        _make_node("S!D8", "=AVERAGE(IF(S!A1:A3>0,S!B1:B3,S!C1:C3))", None),
        _make_node("S!D9", "=MAX(IF(S!A1:A3>0,S!A1:A3))", None),
        _make_node("S!D10", "=MAX(IF(S!A1:A3>0,S!C1:C3))", None),
    )
    result = assert_codegen_matches_evaluator(
        graph, ["S!D1", "S!D2", "S!D3", "S!D4", "S!D5", "S!D6", "S!D7", "S!D8", "S!D9", "S!D10"]
    )
    assert result.generated_results["S!D1"] == 5.0
    assert result.generated_results["S!D2"] == 50.0
    assert result.generated_results["S!D3"] == 150.0
    assert result.generated_results["S!D4"] == 20.0
    assert result.generated_results["S!D5"] == 500.0
    assert result.generated_results["S!D6"] == 50.0
    assert result.generated_results["S!D7"] == 2.5
    assert result.generated_results["S!D8"] == 50.0
    assert result.generated_results["S!D9"] == 3.0
    assert result.generated_results["S!D10"] == 300.0
    assert "xl_if" in result.generated_code
