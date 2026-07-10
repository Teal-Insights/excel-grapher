"""Tests for FormulaEvaluator.evaluate_row (one-row span / full stripe)."""

from __future__ import annotations

import pytest

from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node, make_row_node
from tests.fixtures.row_nodes.option_b_stripe import build_option_b_product_graph


def _wide_option_b_graph() -> DependencyGraph:
    """D63:F63 template `=Sheet1!D35*2` with leaves D35=3, E35=5, F35=7."""
    g = DependencyGraph()
    for col, value in (("D", 3), ("E", 5), ("F", 7)):
        g.add_node(
            Node("Sheet1", col, 35, None, None, value, True),
        )
    row = make_row_node(
        "Sheet1",
        63,
        "D",
        "F",
        formula="=Sheet1!D35*2",
        normalized_formula="=Sheet1!D35*2",
        varying_ref_slots=(0,),
        is_leaf=False,
        is_target=True,
    )
    g.add_node(row)
    for col in ("D", "E", "F"):
        g.add_edge(row.key, f"Sheet1!{col}35")
    return g


def test_evaluate_row_whole_stripe() -> None:
    g = build_option_b_product_graph()
    with FormulaEvaluator(g) as ev:
        assert ev.evaluate_row("Sheet1!D63:E63") == [6, 10]


def test_evaluate_row_subrange() -> None:
    g = _wide_option_b_graph()
    with FormulaEvaluator(g) as ev:
        assert ev.evaluate_row("Sheet1!E63:F63") == [10, 14]
        assert ev.evaluate_row("Sheet1!E63:E63") == [10]


def test_evaluate_row_full_wide_stripe() -> None:
    g = _wide_option_b_graph()
    with FormulaEvaluator(g) as ev:
        assert ev.evaluate_row("Sheet1!D63:F63") == [6, 10, 14]


def test_evaluate_row_is_lazy_outside_span() -> None:
    g = _wide_option_b_graph()
    with FormulaEvaluator(g) as ev:
        assert ev.evaluate_row("Sheet1!E63:E63") == [10]
        assert "Sheet1!E63" in ev._cache
        assert "Sheet1!D63" not in ev._cache
        assert "Sheet1!F63" not in ev._cache


def test_evaluate_row_rejects_single_cell() -> None:
    g = build_option_b_product_graph()
    with FormulaEvaluator(g) as ev, pytest.raises(ValueError, match="one-row"):
        ev.evaluate_row("Sheet1!E63")


def test_evaluate_row_rejects_uncovered_span() -> None:
    g = build_option_b_product_graph()
    with FormulaEvaluator(g) as ev, pytest.raises(KeyError, match="not found"):
        ev.evaluate_row("Sheet1!A1:B1")


def test_evaluate_still_rejects_row_key() -> None:
    """Scalar evaluate() keeps the MVP reject; use evaluate_row for spans."""
    g = build_option_b_product_graph()
    with FormulaEvaluator(g) as ev, pytest.raises(ValueError, match="member"):
        ev.evaluate("Sheet1!D63:E63")
