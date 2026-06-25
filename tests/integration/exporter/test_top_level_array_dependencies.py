"""Integration parity for formulas that consume array-result cells (#284)."""

from __future__ import annotations

import pytest

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


def _array_consumer_graph(formula_address: str, formula: str) -> DependencyGraph:
    graph = DependencyGraph()
    for row, (category, value) in enumerate(
        zip(["Software", "Hardware", "Software"], [10.0, 20.0, 30.0], strict=True),
        start=5,
    ):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
        graph.add_node(_make_node(f"Data!E{row}", None, value))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    graph.add_node(_make_node(formula_address, formula, None))
    return graph


@pytest.mark.parametrize(
    ("target", "formula", "expected"),
    [
        ("Data!E10", "=Data!D10*1", None),
        ("Data!F10", "=SUMPRODUCT(Data!D10,Data!E5:E7)", 40.0),
        ("Data!G10", "=SUM(Data!D10:Data!D12)", 2.0),
        ("Data!H10", "=SUMPRODUCT(Data!D10:Data!D12,Data!E5:E7)", 40.0),
    ],
)
def test_array_consumer_codegen_parity(target: str, formula: str, expected: float | None) -> None:
    graph = _array_consumer_graph(target, formula)
    result = assert_codegen_matches_evaluator(graph, [target])
    if expected is not None:
        assert result.evaluator_results[target] == pytest.approx(expected)
        assert result.generated_results[target] == pytest.approx(expected)
