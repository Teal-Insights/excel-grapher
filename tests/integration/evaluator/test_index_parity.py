"""INDEX edge cases: evaluator and generated export runtime stay aligned (integration).

Uses `assert_codegen_matches_evaluator` for INDEX variants that commonly diverge
between hand-rolled evaluators and transpiled code.
"""

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.types import XlError
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


def test_index_scalar_out_of_bounds_returns_ref_error() -> None:
    """Out-of-bounds scalar INDEX should return #REF! instead of raising."""
    graph = _make_graph(
        _make_node("S!A1", "1", None),
        _make_node("S!B1", "2", None),
        _make_node("S!A2", "3", None),
        _make_node("S!B2", "4", None),
        _make_node("S!C1", "=INDEX(S!A1:S!B2,3,1)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!C1"])
    assert result.evaluator_results["S!C1"] == XlError.REF
    assert result.generated_results["S!C1"] == XlError.REF


def test_index_row_zero_returns_whole_column_for_sum_and_match() -> None:
    """INDEX(array, 0) selects the entire column (Excel whole-axis form; #502)."""
    graph = _make_graph(
        _make_node("S!A1", None, 5),
        _make_node("S!A2", None, 0),
        _make_node("S!A3", None, 7),
        _make_node("S!B1", "=SUM(INDEX(S!A1:S!A3,0))", None),
        _make_node("S!B2", "=MATCH(7,INDEX(S!A1:S!A3,0),0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2"])
    assert result.evaluator_results["S!B1"] == 12
    assert result.generated_results["S!B1"] == 12
    assert result.evaluator_results["S!B2"] == 3
    assert result.generated_results["S!B2"] == 3


def test_index_zero_axis_selectors_on_2d_array() -> None:
    """INDEX(rng, 0, k) / INDEX(rng, k, 0) return whole column/row (#502)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 4),
        _make_node("S!A3", None, 7),
        _make_node("S!B1", None, 2),
        _make_node("S!B2", None, 5),
        _make_node("S!B3", None, 8),
        _make_node("S!C1", None, 3),
        _make_node("S!C2", None, 6),
        _make_node("S!C3", None, 9),
        _make_node("S!E1", "=SUM(INDEX(S!A1:S!C3,0,2))", None),
        _make_node("S!E2", "=SUM(INDEX(S!A1:S!C3,2,0))", None),
        _make_node("S!E3", "=SUM(INDEX(S!A1:S!C3,0,0))", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!E1", "S!E2", "S!E3"])
    assert result.evaluator_results["S!E1"] == 15
    assert result.generated_results["S!E1"] == 15
    assert result.evaluator_results["S!E2"] == 15
    assert result.generated_results["S!E2"] == 15
    assert result.evaluator_results["S!E3"] == 45
    assert result.generated_results["S!E3"] == 45


def test_index_zero_over_computed_array_for_match_true_idiom() -> None:
    """MATCH(TRUE, INDEX((rng<>0),0), 0) finds the first non-zero (#502)."""
    graph = _make_graph(
        _make_node("S!A1", None, 0),
        _make_node("S!A2", None, 0),
        _make_node("S!A3", None, 7),
        _make_node("S!B1", "=MATCH(TRUE,INDEX((S!A1:S!A3<>0),0),0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.evaluator_results["S!B1"] == 3
    assert result.generated_results["S!B1"] == 3
