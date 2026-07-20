"""IF empty/omitted branches: evaluator and codegen agree on Excel 0/FALSE.

Uses `assert_codegen_matches_evaluator` so empty-vs-omitted IF semantics stay
aligned on the evaluator↔codegen path independent of Excel file I/O.
"""

from excel_grapher import DependencyGraph, Node
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


def test_if_empty_vs_omitted_branch_parity() -> None:
    """Empty IF branches are 0; omitted else is FALSE; nested arithmetic stays numeric."""
    graph = _make_graph(
        _make_node("S!A1", "=IF(FALSE,1)", None),  # omitted else -> FALSE
        _make_node("S!A2", "=IF(FALSE,1,)", None),  # empty else -> 0
        _make_node("S!A3", "=IF(TRUE,,5)", None),  # empty then -> 0
        _make_node("S!A4", "=IF(FALSE,,5)", None),  # empty then unused -> 5
        _make_node("S!A5", "=1+IF(FALSE,1,)", None),  # nested empty else -> 1
        _make_node("S!A6", "=ISBLANK(IF(FALSE,1,))", None),  # 0 is not blank
    )

    targets = ["S!A1", "S!A2", "S!A3", "S!A4", "S!A5", "S!A6"]
    result = assert_codegen_matches_evaluator(graph, targets)
    assert result.generated_results["S!A1"] is False
    assert result.generated_results["S!A2"] == 0
    assert result.generated_results["S!A3"] == 0
    assert result.generated_results["S!A4"] == 5
    assert result.generated_results["S!A5"] == 1
    assert result.generated_results["S!A6"] is False
