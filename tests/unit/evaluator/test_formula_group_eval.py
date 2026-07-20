"""Unit tests for formula-group evaluation (Issue 2 sprint 3)."""

from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import AddressHoleNode, AddressLeafKind, CellRefNode
from excel_grapher.evaluator.errors import FormulaGroupKeyError, MissingGroupTemplateError
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.grapher.formula_groups import shape_fingerprint
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node, make_union_node
from tests.fixtures.formula_groups.hand_built import (
    build_cross_sheet_cell_only_twin,
    build_cross_sheet_union_group,
    build_row_stripe_cell_only_twin,
    build_row_stripe_group,
)


def _simple_passthrough_group() -> tuple[DependencyGraph, str]:
    """Group owning A1/B1; skeleton is a single CELL hole bound to Z1/Z2."""
    skeleton = AddressHoleNode(kind=AddressLeafKind.cell, slot=0)
    fp = shape_fingerprint(skeleton)
    group = make_union_node(
        ["Sheet1!A1", "Sheet1!B1"],
        is_leaf=False,
        shape_fingerprint=fp,
        skeleton=skeleton,
        member_bindings={
            "Sheet1!A1": (CellRefNode(address="Sheet1!Z1"),),
            "Sheet1!B1": (CellRefNode(address="Sheet1!Z2"),),
        },
    )
    g = DependencyGraph()
    g.add_node(make_cell_node("Sheet1", "Z", 1, value=10.0))
    g.add_node(make_cell_node("Sheet1", "Z", 2, value=20.0))
    g.add_node(group)
    g.add_edge(group.key, "Sheet1!Z1")
    g.add_edge(group.key, "Sheet1!Z2")
    return g, group.key


def _simple_cell_only_twin() -> DependencyGraph:
    g = DependencyGraph()
    g.add_node(make_cell_node("Sheet1", "Z", 1, value=10.0))
    g.add_node(make_cell_node("Sheet1", "Z", 2, value=20.0))
    g.add_node(
        make_cell_node(
            "Sheet1",
            "A",
            1,
            formula="=Sheet1!Z1",
            normalized_formula="=Sheet1!Z1",
            is_leaf=False,
        )
    )
    g.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            1,
            formula="=Sheet1!Z2",
            normalized_formula="=Sheet1!Z2",
            is_leaf=False,
        )
    )
    g.add_edge("Sheet1!A1", "Sheet1!Z1")
    g.add_edge("Sheet1!B1", "Sheet1!Z2")
    return g


def test_evaluate_member_matches_cell_only_twin() -> None:
    group_graph, _group_key = _simple_passthrough_group()
    twin = _simple_cell_only_twin()
    with FormulaEvaluator(group_graph) as ev_g, FormulaEvaluator(twin) as ev_t:
        assert ev_g.evaluate("Sheet1!A1") == ev_t.evaluate("Sheet1!A1") == 10.0
        assert ev_g.evaluate("Sheet1!B1") == ev_t.evaluate("Sheet1!B1") == 20.0


def test_evaluate_member_is_lazy_for_siblings() -> None:
    g, _ = _simple_passthrough_group()
    with FormulaEvaluator(g) as ev:
        assert ev.evaluate("Sheet1!A1") == 10.0
        assert "Sheet1!A1" in ev._cache
        assert "Sheet1!B1" not in ev._cache
        assert ev.evaluate("Sheet1!B1") == 20.0
        assert "Sheet1!B1" in ev._cache


def test_evaluate_rejects_group_key() -> None:
    g, group_key = _simple_passthrough_group()
    with FormulaEvaluator(g) as ev, pytest.raises(FormulaGroupKeyError, match="multi-cell"):
        ev.evaluate(group_key)


def test_evaluate_rejects_union_key_string() -> None:
    fx = build_cross_sheet_union_group()
    with (
        FormulaEvaluator(fx.graph) as ev,
        pytest.raises(FormulaGroupKeyError, match="multi-cell"),
    ):
        ev.evaluate(fx.group_key)


def test_evaluate_missing_template_raises() -> None:
    g = DependencyGraph()
    g.add_node(make_cell_node("Sheet1", "Z", 1, value=1.0))
    # Multi-cell node without template fields.
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!B1"], is_leaf=False))
    with (
        FormulaEvaluator(g) as ev,
        pytest.raises(MissingGroupTemplateError, match="no usable template"),
    ):
        ev.evaluate("Sheet1!A1")


def test_cell_only_path_unchanged() -> None:
    twin = _simple_cell_only_twin()
    with FormulaEvaluator(twin) as ev:
        assert ev.evaluate("Sheet1!A1") == 10.0
        assert ev.evaluate("Sheet1!Z1") == 10.0


def test_row_stripe_group_eval_matches_twin() -> None:
    fx = build_row_stripe_group()
    twin = build_row_stripe_cell_only_twin()
    with FormulaEvaluator(fx.graph) as ev_g, FormulaEvaluator(twin) as ev_t:
        with pytest.raises(FormulaGroupKeyError):
            ev_g.evaluate(fx.group_key)
        for member in fx.members:
            assert ev_g.evaluate(member) == ev_t.evaluate(member)
        # Laziness across members
        ev_g.clear_caches()
        _ = ev_g.evaluate("Sheet1!E63")
        assert "Sheet1!E63" in ev_g._cache
        assert "Sheet1!D63" not in ev_g._cache
        assert "Sheet1!F63" not in ev_g._cache


def test_cross_sheet_group_eval_matches_twin() -> None:
    fx = build_cross_sheet_union_group()
    twin = build_cross_sheet_cell_only_twin()
    with FormulaEvaluator(fx.graph) as ev_g, FormulaEvaluator(twin) as ev_t:
        with pytest.raises(FormulaGroupKeyError):
            ev_g.evaluate(fx.group_key)
        for member in fx.members:
            assert fx.graph.get_node(member) is None
            assert ev_g.evaluate(member) == ev_t.evaluate(member)
