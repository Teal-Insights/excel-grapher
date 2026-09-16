"""VALUE: evaluator and generated export runtime agree on synthetic graphs (integration).

Guards numeric parsing parity through `evaluate_targets` for small
dependency graphs.
"""

from datetime import datetime

import pytest

from excel_grapher import DependencyGraph, Node, XlError
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.coercions import datetime_to_excel_serial
from tests.integration.utils.parity_harness import evaluate_targets


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


def test_value_parity() -> None:
    """VALUE parses numbers, currency, ISO dates, blanks, and error/invalid text."""
    expected_date_serial = datetime_to_excel_serial(datetime(2018, 3, 15))
    graph = _make_graph(
        _make_node("S!A1", '=VALUE("42")', None),
        _make_node("S!A2", '=VALUE("12.5")', None),
        _make_node("S!A3", '=VALUE("6522014")', None),
        _make_node("S!A4", '=VALUE("$1,234.56")', None),
        _make_node("S!A5", '=VALUE("abc")', None),
        _make_node("S!A6", "=VALUE(#REF!)", None),
        _make_node("S!A7", '=VALUE("")', None),
        _make_node("S!A8", '=VALUE("2018-03-15")', None),
    )

    results = evaluate_targets(graph, [f"S!A{i}" for i in range(1, 9)])
    assert results["S!A1"] == 42.0
    assert results["S!A2"] == 12.5
    assert results["S!A3"] == 6522014.0
    assert results["S!A4"] == 1234.56
    assert results["S!A5"] == XlError.VALUE
    assert results["S!A6"] == XlError.REF
    assert results["S!A7"] == 0.0
    assert results["S!A8"] == pytest.approx(expected_date_serial)
