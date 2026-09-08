"""Missing cell references raise in the evaluator (integration).

Formulas that reference cells absent from the graph fail closed with `KeyError`.
"""

import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.core.address_keys import parse_address


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


def test_missing_cell_reference_raises() -> None:
    """Missing cell access should raise in the evaluator."""
    graph = _make_graph(
        _make_node("S!A1", "=S!B1+1", None),
        # S!B1 is NOT in the graph
    )

    with pytest.raises(KeyError, match="S!B1"), FormulaEvaluator(graph) as ev:
        ev.evaluate(["S!A1"])
