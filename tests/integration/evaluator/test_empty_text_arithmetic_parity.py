"""Empty text vs blank arithmetic: evaluator ↔ export parity (#420).

Guard formulas often return empty text. Excel raises `#VALUE!` for arithmetic
on empty text, while blank cells still coerce to 0 (`#DIV/0!` for division).
"""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node, XlError
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


def test_empty_text_arithmetic_parity_matches_excel_error_classes() -> None:
    """B1/B3: empty text -> #VALUE!; B2: blank cell divisor -> #DIV/0!."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", '=IF(TRUE,"",0)', None),
        # Blank leaf (`None`) still coerces to 0 in arithmetic.
        _make_node("S!A3", None, None),
        _make_node("S!B1", "=S!A1/S!A2", None),
        _make_node("S!B2", "=S!A1/S!A3", None),
        _make_node("S!B3", "=S!A1+S!A2", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2", "S!B3"])
    assert result.evaluator_results["S!B1"] == XlError.VALUE
    assert result.evaluator_results["S!B2"] == XlError.DIV
    assert result.evaluator_results["S!B3"] == XlError.VALUE


def test_sum_skips_empty_text_while_arithmetic_rejects_it() -> None:
    """SUM ignores guard empty text; `+` still raises `#VALUE!` (parity)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", '=IF(TRUE,"",0)', None),
        _make_node("S!A3", None, 2),
        _make_node("S!B1", "=SUM(S!A1:S!A3)", None),
        _make_node("S!B2", "=S!A1+S!A2", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2"])
    assert result.evaluator_results["S!B1"] == 3.0
    assert result.evaluator_results["S!B2"] == XlError.VALUE
