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
    rewrite_guard_aliases,
)
from excel_grapher.grapher.may_cycle import GuardConeAbstraction, identity_alias_map
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


# ---- affine / interval abstraction of guard-cone cells (#1028) ---------------


def _cycle_with_guard_cells(guard, cells: dict[str, str], env=None):
    """Two-node guarded cycle whose guard cells have the given formulas."""
    nodes = [
        make_cell_node("Sheet1", key[0], int(key[1:]), normalized_formula=f, is_leaf=False)
        for key, f in cells.items()
    ]
    graph = _two_node_guarded_cycle(a_to_b=guard, b_to_a=guard, extra_nodes=nodes)
    return graph.cycle_report(cell_type_env=env)


def _cmp(left, op: str, right):
    def term(x):
        return Literal(x) if isinstance(x, (int, float)) else CellRef(f"Sheet1!{x}")

    return Compare(left=term(left), op=op, right=term(right))


def test_same_root_offsets_decide_cell_cell_equality() -> None:
    cells = {"P1": "=Sheet1!L1+1", "Q1": "=Sheet1!L1+2"}
    assert _cycle_with_guard_cells(_cmp("P1", "=", "Q1"), cells).has_may_cycles is False
    assert _cycle_with_guard_cells(_cmp("P1", "<>", "Q1"), cells).has_may_cycles is True


def test_offset_chains_fold_to_a_common_root() -> None:
    cells = {"P1": "=Sheet1!L1+1", "P2": "=Sheet1!P1+1", "Q1": "=Sheet1!L1+2", "Q2": "=Sheet1!Q1"}
    assert _cycle_with_guard_cells(_cmp("P2", "=", "Q2"), cells).has_may_cycles is True
    assert _cycle_with_guard_cells(_cmp("P2", "<", "Q2"), cells).has_may_cycles is False


def test_max_index_outside_its_interval_is_infeasible() -> None:
    env = constraints_to_cell_type_env(
        {"Sheet1!L1": TypingLiteral[0, 1, 2, 3], "Sheet1!K1": TypingLiteral[0, 1, 2, 3]}, {}
    )
    for formula in (
        "=MAX(Sheet1!L1-Sheet1!K1,0)",
        "=IF(Sheet1!L1-Sheet1!K1>0,Sheet1!L1-Sheet1!K1,0)",
    ):
        cells = {"X1": formula}
        assert _cycle_with_guard_cells(_cmp("X1", "=", 5), cells, env).has_may_cycles is False
        assert _cycle_with_guard_cells(_cmp("X1", "<", 0), cells, env).has_may_cycles is False
        assert _cycle_with_guard_cells(_cmp("X1", "=", 2), cells, env).has_may_cycles is True


def test_max_of_same_root_difference_is_constant() -> None:
    cells = {
        "A2": "=Sheet1!L1+5",
        "B2": "=Sheet1!L1+2",
        "X1": "=IF(Sheet1!A2-Sheet1!B2>0,Sheet1!A2-Sheet1!B2,0)",
    }
    x_is_zero = _cmp("X1", "=", 0)
    assert _cycle_with_guard_cells(x_is_zero, cells).has_may_cycles is False
    assert _cycle_with_guard_cells(Not(x_is_zero), cells).has_may_cycles is True
    assert _cycle_with_guard_cells(_cmp("X1", "=", 4), cells).has_may_cycles is False


def test_unknown_formula_shape_keeps_edge_feasible() -> None:
    env = constraints_to_cell_type_env({"Sheet1!L1": TypingLiteral[0, 1]}, {})
    cells = {"X1": "=Sheet1!L1*7"}
    assert _cycle_with_guard_cells(_cmp("X1", "=", 7), cells, env).has_may_cycles is True
    assert _cycle_with_guard_cells(_cmp("X1", "=", 99), cells, env).has_may_cycles is True


def test_ordered_compare_of_possibly_boolean_copy_is_not_decided() -> None:
    """With L1=TRUE, Excel ranks TRUE above 2, so `P1>=Q1` can hold."""
    cells = {"P1": "=Sheet1!L1", "Q1": "=Sheet1!L1+1"}
    assert _cycle_with_guard_cells(_cmp("P1", ">=", "Q1"), cells).has_may_cycles is True
    assert _cycle_with_guard_cells(_cmp("P1", "=", "Q1"), cells).has_may_cycles is False


def test_abstraction_handles_long_chains_and_formula_cycles() -> None:
    nodes = {
        f"S!A{i}": make_cell_node("S", "A", i, normalized_formula=f"=S!A{i - 1}+1", is_leaf=False)
        for i in range(2, 5002)
    }
    nodes["S!B1"] = make_cell_node("S", "B", 1, normalized_formula="=S!B2+1", is_leaf=False)
    nodes["S!B2"] = make_cell_node("S", "B", 2, normalized_formula="=S!B1-1", is_leaf=False)
    cone = GuardConeAbstraction(nodes, None)
    tail = cone.value("S!A5001")
    assert (tail.root, tail.offset) == ("S!A1", 5000.0)
    assert cone.simplify(Compare(left=CellRef("S!A5001"), op="=", right=CellRef("S!A2"))) is False
    guard = Compare(left=CellRef("S!B1"), op="=", right=Literal(3))
    assert cone.simplify(guard) is guard


def _guarded_ring(layers: int) -> DependencyGraph:
    """Ring of two-way IF layers closed by a back edge guarded by `Sheet1!P1=0`.

    Every layer picks one of two guarded paths, so a path-sensitive search that
    cannot refute the back edge up front visits about `2**layers` states.
    """
    graph = DependencyGraph()
    for j in range(1, layers + 1):
        for col in "CDE":
            graph.add_node(make_cell_node("Sheet1", col, j, formula="=0", is_leaf=False))
    for j in range(1, layers + 1):
        sel = CellRef(f"Sheet1!S{j}")
        graph.add_edge(f"Sheet1!C{j}", f"Sheet1!D{j}", guard=Compare(sel, "=", Literal(1)))
        graph.add_edge(f"Sheet1!C{j}", f"Sheet1!E{j}", guard=Not(Compare(sel, "=", Literal(1))))
        nxt = f"Sheet1!C{j + 1}" if j < layers else "Sheet1!C1"
        back = Compare(CellRef("Sheet1!P1"), "=", Literal(0)) if j == layers else None
        graph.add_edge(f"Sheet1!D{j}", nxt, guard=back)
        graph.add_edge(f"Sheet1!E{j}", nxt, guard=back)
    return graph


def _ring_env(layers: int, pins: int):
    constraints: dict[str, object] = {"Sheet1!P1": TypingLiteral[1]}
    constraints |= {f"Sheet1!S{j}": TypingLiteral[1, 2] for j in range(1, layers + 1)}
    constraints |= {f"Sheet1!Z{i}": TypingLiteral[0] for i in range(1, pins + 1)}
    return constraints_to_cell_type_env(constraints, {})


def test_irrelevant_singleton_pins_do_not_enter_guard_constraints(monkeypatch) -> None:
    """#1042: pins on cells no guard mentions must not be seeded into DFS states."""
    from excel_grapher.grapher import guard as guard_mod

    largest = 0
    real_add = guard_mod.GuardConstraints.add

    def spy(self, g, *, cell_type_env=None):
        nonlocal largest
        largest = max(largest, len(self.equalities))
        return real_add(self, g, cell_type_env=cell_type_env)

    monkeypatch.setattr(guard_mod.GuardConstraints, "add", spy)
    graph = _guarded_ring(4)
    report = graph.cycle_report(cell_type_env=_ring_env(4, pins=500))
    assert report.has_may_cycles is False
    assert largest <= 5


def test_edge_refuted_by_env_alone_eliminates_scc_without_path_search(monkeypatch) -> None:
    """#1042: an SCC broken by edge-local refutation needs no path-sensitive DFS."""
    from excel_grapher.grapher import graph as graph_mod

    def boom(*args, **kwargs):
        raise AssertionError("path search should not run")

    monkeypatch.setattr(graph_mod, "_subgraph_has_feasible_cycle", boom)
    graph = _guarded_ring(12)
    report = graph.cycle_report(cell_type_env=_ring_env(12, pins=0))
    assert report.has_may_cycles is False


def test_edge_refutation_keeps_reporting_the_whole_scc() -> None:
    """Refuting one edge that leaves an inner cycle still reports the original SCC."""
    graph = _guarded_ring(2)
    # Inner loop C1 -> D1 -> C1, feasible, alongside the refutable ring.
    graph.add_edge("Sheet1!D1", "Sheet1!C1", guard=None)
    report = graph.cycle_report(cell_type_env=_ring_env(2, pins=0))
    assert report.has_may_cycles is True
    assert len(report.may_cycles) == 1
    assert "Sheet1!C2" in report.may_cycles[0]
