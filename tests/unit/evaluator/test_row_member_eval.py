"""Evaluator dispatch for Option B row members (issue #377 sprint 3)."""

from __future__ import annotations

import pytest

from excel_grapher.evaluator.evaluator import FormulaEvaluator
from tests.fixtures.row_nodes.option_b_stripe import (
    OPTION_B_ROW_KEY,
    build_option_b_stripe_fixture,
)


def test_eval_member_via_locate_matches_cell_only_twin() -> None:
    fixture = build_option_b_stripe_fixture()
    with (
        FormulaEvaluator(fixture.option_b) as row_ev,
        FormulaEvaluator(fixture.cell_only) as twin_ev,
    ):
        for member in fixture.member_keys:
            assert row_ev.evaluate(member) == twin_ev.evaluate(member)


def test_eval_all_members_match_expected_scalars() -> None:
    fixture = build_option_b_stripe_fixture()
    with FormulaEvaluator(fixture.option_b) as ev:
        assert ev.evaluate("Sheet1!D63") == 6
        assert ev.evaluate("Sheet1!E63") == 10


def test_eval_is_lazy_per_member() -> None:
    fixture = build_option_b_stripe_fixture()
    with FormulaEvaluator(fixture.option_b) as ev:
        assert ev.evaluate("Sheet1!E63") == 10
        assert "Sheet1!E63" in ev._cache
        assert "Sheet1!D63" not in ev._cache


def test_eval_rejects_row_key_mvp() -> None:
    fixture = build_option_b_stripe_fixture()
    with FormulaEvaluator(fixture.option_b) as ev, pytest.raises(ValueError, match="member"):
        ev.evaluate(OPTION_B_ROW_KEY)


def test_eval_member_with_abs_row_markers_in_template() -> None:
    """Templates may keep Excel `$` markers; eval normalizes before parse."""
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.node import Node, make_row_node

    g = DependencyGraph()
    g.add_node(
        Node("Sheet1", "D", 35, None, None, 3, True),
    )
    g.add_node(
        Node("Sheet1", "E", 35, None, None, 5, True),
    )
    g.add_node(
        make_row_node(
            "Sheet1",
            63,
            "D",
            "E",
            formula="=Sheet1!D$35*2",
            normalized_formula="=Sheet1!D$35*2",
            varying_ref_slots=(0,),
            is_leaf=False,
        )
    )
    with FormulaEvaluator(g) as ev:
        assert ev.evaluate("Sheet1!E63") == 10
