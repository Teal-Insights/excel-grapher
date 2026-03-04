from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.name_utils import parse_address
from tests.evaluator.parity_harness import (
    assert_code_does_not_embed_symbols,
    assert_codegen_matches_evaluator,
)


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


def test_golden_parity_mixed_features_small_graph() -> None:
    """A compact scenario mixing coercions, comparisons, IF/IFERROR, and ranges."""
    graph = _make_graph(
        _make_node("S!A1", None, "0"),  # numeric string
        _make_node("S!A2", None, 0),
        _make_node("S!A3", None, "FALSE"),
        _make_node("S!B1", "=S!A1=S!A2", None),  # True via numeric-string coercion
        _make_node("S!B2", "=IF(S!A3,1,2)", None),  # 2 via Excel boolean coercion
        _make_node("S!B3", "=IFERROR(1/0,99)", None),  # 99 (DIV/0! caught)
        _make_node("S!B4", "=SUM(S!B2,S!B3)", None),  # 101
        _make_node("S!C1", "=SUM(S!B1:S!B4)", None),  # includes boolean -> 1 + 2 + 99 + 101 = 203
    )

    targets = ["S!B1", "S!B2", "S!B3", "S!B4", "S!C1"]
    result = assert_codegen_matches_evaluator(graph, targets)

    assert result.generated_results["S!B1"] is True
    assert result.generated_results["S!B2"] == 2
    assert result.generated_results["S!B3"] == 99
    assert result.generated_results["S!B4"] == 101.0
    assert result.generated_results["S!C1"] == 203.0

    # No OFFSET in this scenario, so dynamic OFFSET runtime should not be included.
    assert "_CELL_TABLE = {" not in result.generated_code
    assert_code_does_not_embed_symbols(result.generated_code, absent={"xl_offset"})

