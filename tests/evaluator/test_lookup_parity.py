"""Parity tests for LOOKUP."""

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


def test_vlookup_parity_exact_and_approximate() -> None:
    graph = _make_graph(
        # Table: A1:B5
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!A3", None, 30),
        _make_node("S!A4", None, 40),
        _make_node("S!A5", None, 50),
        _make_node("S!B1", None, "ten"),
        _make_node("S!B2", None, "twenty"),
        _make_node("S!B3", None, "thirty"),
        _make_node("S!B4", None, "forty"),
        _make_node("S!B5", None, "fifty"),
        # Exact match
        _make_node("S!C1", "=VLOOKUP(20, S!A1:B5, 2, FALSE)", None),
        _make_node("S!C2", "=VLOOKUP(25, S!A1:B5, 2, FALSE)", None),
        # Approximate match (range_lookup=TRUE)
        _make_node("S!C3", "=VLOOKUP(25, S!A1:B5, 2, TRUE)", None),
        _make_node("S!C4", "=VLOOKUP(5, S!A1:B5, 2, TRUE)", None),
        _make_node("S!C5", "=VLOOKUP(100, S!A1:B5, 2, TRUE)", None),
        # Return column 1
        _make_node("S!C6", "=VLOOKUP(30, S!A1:B5, 1, FALSE)", None),
        # Case-insensitive string lookup
        _make_node("S!D1", None, "Alpha"),
        _make_node("S!D2", None, "Beta"),
        _make_node("S!D3", None, "Gamma"),
        _make_node("S!E1", None, 1),
        _make_node("S!E2", None, 2),
        _make_node("S!E3", None, 3),
        _make_node("S!F1", '=VLOOKUP("alpha", S!D1:E3, 2, FALSE)', None),
    )

    result = assert_codegen_matches_evaluator(
        graph, ["S!C1", "S!C2", "S!C3", "S!C4", "S!C5", "S!C6", "S!F1"]
    )
    assert result.generated_results["S!C1"] == "twenty"
    assert result.generated_results["S!C2"] == XlError.NA
    assert result.generated_results["S!C3"] == "twenty"
    assert result.generated_results["S!C4"] == XlError.NA
    assert result.generated_results["S!C5"] == "fifty"
    assert result.generated_results["S!C6"] == 30
    assert result.generated_results["S!F1"] == 1


def test_lookup_parity_vector_and_array_forms() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!B3", None, 30),
        _make_node("S!C1", "=LOOKUP(2.5, S!A1:A3, S!B1:B3)", None),
        _make_node("S!C2", "=LOOKUP(0.5, S!A1:A3, S!B1:B3)", None),
        _make_node("S!C3", "=LOOKUP(5, S!A1:A3, S!B1:B3)", None),
        _make_node("S!C4", "=LOOKUP(2, S!A1:A3)", None),
        _make_node("S!D1", None, 1),
        _make_node("S!D2", None, 2),
        _make_node("S!D3", None, 3),
        _make_node("S!E1", None, 10),
        _make_node("S!E2", None, 20),
        _make_node("S!E3", None, 30),
        _make_node("S!F1", "=LOOKUP(2.5, S!D1:E3)", None),
        _make_node("S!G1", None, 1),
        _make_node("S!H1", None, 2),
        _make_node("S!I1", None, 3),
        _make_node("S!G2", None, 10),
        _make_node("S!H2", None, 20),
        _make_node("S!I2", None, 30),
        _make_node("S!J1", "=LOOKUP(2.5, S!G1:I2)", None),
    )

    result = assert_codegen_matches_evaluator(
        graph, ["S!C1", "S!C2", "S!C3", "S!C4", "S!F1", "S!J1"]
    )
    assert result.generated_results["S!C1"] == 20
    assert result.generated_results["S!C2"] == XlError.NA
    assert result.generated_results["S!C3"] == 30
    assert result.generated_results["S!C4"] == 2
    assert result.generated_results["S!F1"] == 20
    assert result.generated_results["S!J1"] == 20
