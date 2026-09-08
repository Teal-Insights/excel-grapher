"""May-cycle feasibility: cell-cell guards, leaf domains, identity aliases (#533)."""

from __future__ import annotations

from typing import Literal as TypingLiteral

from excel_grapher.core.cell_types import constraints_to_cell_type_env
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import (
    CellRef,
    Compare,
    Literal,
    Not,
    evaluate_guard,
    rewrite_guard_aliases,
)
from excel_grapher.grapher.may_cycle import identity_alias_map
from excel_grapher.grapher.node import make_cell_node


def _two_node_guarded_cycle(
    *,
    a_to_b,
    b_to_a,
    extra_nodes: list | None = None,
) -> DependencyGraph:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "A", 1, formula="=IF(1,B1,0)", is_leaf=False))
    graph.add_node(make_cell_node("Sheet1", "B", 1, formula="=IF(1,A1,0)", is_leaf=False))
    graph.add_edge("Sheet1!A1", "Sheet1!B1", guard=a_to_b)
    graph.add_edge("Sheet1!B1", "Sheet1!A1", guard=b_to_a)
    for node in extra_nodes or []:
        graph.add_node(node)
    return graph


def test_cell_cell_equal_and_unequal_on_opposite_edges_is_infeasible() -> None:
    eq = Compare(left=CellRef("Sheet1!X1"), op="=", right=CellRef("Sheet1!Y1"))
    graph = _two_node_guarded_cycle(a_to_b=eq, b_to_a=Not(eq))
    report = graph.cycle_report()
    assert report.has_must_cycles is False
    assert report.has_may_cycles is False


def test_same_cell_cell_equality_on_both_edges_stays_feasible() -> None:
    eq = Compare(left=CellRef("Sheet1!X1"), op="=", right=CellRef("Sheet1!Y1"))
    graph = _two_node_guarded_cycle(a_to_b=eq, b_to_a=eq)
    report = graph.cycle_report()
    assert report.has_may_cycles is True


def test_leaf_enum_domain_makes_complementary_if_guards_infeasible() -> None:
    """Sheet1!B8 in {0,1} makes NOT(B8=0) AND NOT(B8=1) unsatisfiable."""
    graph = _two_node_guarded_cycle(
        a_to_b=Not(Compare(left=CellRef("Sheet1!B8"), op="=", right=Literal(0))),
        b_to_a=Not(Compare(left=CellRef("Sheet1!B8"), op="=", right=Literal(1))),
        extra_nodes=[make_cell_node("Sheet1", "B", 8, value=0, is_leaf=True)],
    )
    env = constraints_to_cell_type_env({"Sheet1!B8": TypingLiteral[0, 1]}, {})
    without = graph.cycle_report()
    with_arg = graph.cycle_report(cell_type_env=env)
    graph.cell_type_env = env
    with_attr = graph.cycle_report()
    assert without.has_may_cycles is True
    assert with_arg.has_may_cycles is False
    assert with_attr.has_may_cycles is False


def test_identity_alias_rewrites_guard_cell_to_leaf() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "L", 1, value="Residency-based", is_leaf=True))
    graph.add_node(
        make_cell_node("Sheet1", "A", 1, normalized_formula="=Sheet1!L1", is_leaf=False),
    )
    aliases = identity_alias_map(graph._nodes)
    assert aliases["Sheet1!A1"] == "Sheet1!L1"
    guard = Compare(left=CellRef("Sheet1!A1"), op="=", right=Literal("Residency-based"))
    rewritten = rewrite_guard_aliases(guard, aliases)
    assert rewritten == Compare(left=CellRef("Sheet1!L1"), op="=", right=Literal("Residency-based"))


def test_identity_alias_plus_singleton_domain_kills_mismatched_equality() -> None:
    """Alias = Const is unsat when the identity leaf is a different singleton."""
    graph = DependencyGraph()
    graph.add_node(
        make_cell_node("Lookup", "X", 4, normalized_formula="=Translation!C90", is_leaf=False)
    )
    graph.add_node(make_cell_node("Translation", "C", 90, value="Residency-based", is_leaf=True))
    graph.add_node(make_cell_node("Sheet1", "A", 1, formula="=IF(1,B1,0)", is_leaf=False))
    graph.add_node(make_cell_node("Sheet1", "B", 1, formula="=IF(1,A1,0)", is_leaf=False))
    mismatch = Compare(
        left=CellRef("Lookup!X4"),
        op="=",
        right=Literal("Currency-based"),
    )
    graph.add_edge("Sheet1!A1", "Sheet1!B1", guard=mismatch)
    graph.add_edge("Sheet1!B1", "Sheet1!A1", guard=mismatch)
    env = constraints_to_cell_type_env(
        {"Translation!C90": TypingLiteral["Residency-based"]},
        {},
    )
    assert graph.cycle_report().has_may_cycles is True
    assert graph.cycle_report(cell_type_env=env).has_may_cycles is False


def test_evaluate_guard_three_valued_and() -> None:
    g = Compare(left=CellRef("Sheet1!A1"), op="=", right=Literal(0))
    assert evaluate_guard(g, {"Sheet1!A1": 0}) is True
    assert evaluate_guard(g, {"Sheet1!A1": 1}) is False
    assert evaluate_guard(g, {}) is None
    assert evaluate_guard(None, {}) is True
