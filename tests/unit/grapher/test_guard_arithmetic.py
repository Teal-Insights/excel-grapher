"""Arithmetic guard operands and ordered comparisons (difference bounds, #1029)."""

from __future__ import annotations

import pickle
from typing import Annotated
from typing import Literal as TypingLiteral

import pytest

from excel_grapher.core.cell_types import Between, RealBetween, constraints_to_cell_type_env
from excel_grapher.grapher.cache import _guard_from_json, _guard_to_json
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import (
    And,
    Arith,
    CellRef,
    Compare,
    GuardConstraints,
    GuardExpr,
    Literal,
    Neg,
    Not,
    instantiate_element_guard,
    rewrite_guard_aliases,
    rewrite_guard_keys,
)
from excel_grapher.grapher.node import make_cell_node
from excel_grapher.grapher.parser import parse_guard_expr

A, B, C, D = (CellRef(f"S!{c}1") for c in "ABCD")


def _cmp(left: GuardExpr, op: str, right: GuardExpr) -> Compare:
    return Compare(left=left, op=op, right=right)


def _add(*guards: GuardExpr) -> GuardConstraints | None:
    out: GuardConstraints | None = GuardConstraints()
    for g in guards:
        assert out is not None
        out = out.add(g)
    return out


# --- Parser -----------------------------------------------------------------


def test_parse_ordered_compare_with_sum_operands_inside_and() -> None:
    parsed = parse_guard_expr("AND(S!A1>S!B1+S!C1,S!A1<=S!B1+S!D1)", current_sheet="S")
    assert parsed == And(
        (
            _cmp(A, ">", Arith(B, "+", C)),
            _cmp(A, "<=", Arith(B, "+", D)),
        )
    )


def test_parse_difference_operand() -> None:
    assert parse_guard_expr("S!A1-S!B1>0", current_sheet="S") == _cmp(
        Arith(A, "-", B), ">", Literal(0)
    )


def test_parse_arithmetic_is_left_associative_and_keeps_parentheses() -> None:
    assert parse_guard_expr("A1-B1+2>=(C1-D1)", current_sheet="S") == _cmp(
        Arith(Arith(A, "-", B), "+", Literal(2)), ">=", Arith(C, "-", D)
    )


def test_parse_unary_minus_on_cell() -> None:
    assert parse_guard_expr("-A1<B1", current_sheet="S") == _cmp(Neg(A), "<", B)


def test_parse_negative_literal_stays_literal() -> None:
    assert parse_guard_expr("A1>-3", current_sheet="S") == _cmp(A, ">", Literal(-3))


@pytest.mark.parametrize(
    "expr",
    [
        'A1>B1+"x"',  # text in arithmetic is #VALUE!
        "A1>B1*C1",  # only + and - are modelled
        "A1>B1+",
        "A1+B1",  # arithmetic only inside a comparison
    ],
)
def test_parse_rejects_unsupported_arithmetic(expr: str) -> None:
    assert parse_guard_expr(expr, current_sheet="S") is None


def test_guard_str_round_trips_through_parser() -> None:
    g = _cmp(Arith(A, "-", Arith(B, "+", C)), ">", Neg(D))
    assert parse_guard_expr(str(g), current_sheet="S") == g


# --- Guard tree walkers -------------------------------------------------------


def test_rewrites_reach_cells_inside_arithmetic() -> None:
    g = _cmp(A, ">", Arith(B, "+", Neg(C)))
    assert rewrite_guard_keys(g, "S!C1", "S!D1") == _cmp(A, ">", Arith(B, "+", Neg(D)))
    assert rewrite_guard_aliases(g, {"S!B1": "S!D1"}) == _cmp(A, ">", Arith(D, "+", Neg(C)))


def test_instantiate_element_guard_through_arithmetic() -> None:
    from excel_grapher.grapher.guard import RangeRef

    template = _cmp(RangeRef("S!A1:A3"), ">", Arith(B, "+", Literal(1)))
    assert instantiate_element_guard(template, row_offset=2, col_offset=0) == _cmp(
        CellRef("S!A3"), ">", Arith(B, "+", Literal(1))
    )


def test_arithmetic_guard_round_trips_through_cache_json_and_pickle() -> None:
    g = And((_cmp(A, ">", Arith(B, "+", C)), _cmp(Neg(A), "<=", Arith(B, "-", Literal(2.5)))))
    assert _guard_from_json(_guard_to_json(g)) == g

    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "E", 1, formula="=IF(1,F1,0)", is_leaf=False))
    graph.add_node(make_cell_node("S", "F", 1, value=1, is_leaf=True))
    graph.add_edge("S!E1", "S!F1", guard=g)
    restored = pickle.loads(pickle.dumps(graph))
    assert restored.get_edge_guard("S!E1", "S!F1") == g


# --- Difference-bound solver --------------------------------------------------


def test_equal_then_greater_is_infeasible() -> None:
    assert _add(_cmp(A, "=", B), _cmp(A, ">", B)) is None
    assert _add(_cmp(A, ">", B), _cmp(B, "=", A)) is None


def test_non_strict_both_ways_is_feasible_but_strict_is_not() -> None:
    assert _add(_cmp(A, ">=", B), _cmp(B, ">=", A)) is not None
    assert _add(_cmp(A, ">=", B), _cmp(B, ">", A)) is None


def test_transitive_strict_chain_is_infeasible() -> None:
    assert _add(_cmp(A, ">", B), _cmp(B, ">", C)) is not None
    assert _add(_cmp(A, ">", B), _cmp(B, ">", C), _cmp(C, ">", A)) is None


def test_offsets_on_both_sides() -> None:
    # A - B > 3 and B > A - 4  <=>  3 < A - B < 4
    assert (
        _add(_cmp(Arith(A, "-", B), ">", Literal(3)), _cmp(B, ">", Arith(A, "-", Literal(4))))
        is not None
    )
    assert (
        _add(_cmp(Arith(A, "-", B), ">", Literal(3)), _cmp(B, ">=", Arith(A, "-", Literal(3))))
        is None
    )


def test_not_of_ordered_compare_is_its_complement() -> None:
    assert _add(_cmp(A, "<", Literal(5)), Not(_cmp(A, "<", Literal(5)))) is None
    assert _add(_cmp(A, "<=", Literal(5)), Not(_cmp(A, "<", Literal(5)))) is not None


def test_literal_equality_meets_ordered_bounds() -> None:
    assert _add(_cmp(A, "=", Literal(3)), _cmp(A, ">", Literal(4))) is None
    assert _add(_cmp(A, "=", Literal(3)), _cmp(A, "=", B), _cmp(B, "<", Literal(3))) is None
    assert _add(_cmp(A, "=", Literal(3)), _cmp(A, "<", Arith(B, "+", Literal(1)))) is not None


def test_equalities_after_ordered_bounds_still_conflict() -> None:
    """Equalities join the bounds lazily, whichever order the guards arrive in."""
    assert _add(_cmp(A, ">", Literal(4)), _cmp(A, "=", Literal(3))) is None
    assert _add(_cmp(B, "<", Literal(3)), _cmp(A, "=", Literal(3)), _cmp(A, "=", B)) is None
    assert _add(_cmp(A, ">", C), _cmp(B, ">", A), _cmp(C, "=", B)) is None
    assert _add(_cmp(A, ">", Literal(4)), _cmp(A, "=", Literal(5))) is not None


def test_plain_equalities_do_not_populate_bounds() -> None:
    c = _add(_cmp(A, "=", Literal(3)), _cmp(A, "=", B), Not(_cmp(C, "=", Literal(1))))
    assert c is not None
    assert c.bounds == ()


def test_equality_with_arithmetic_gives_two_bounds() -> None:
    assert _add(_cmp(A, "=", Arith(B, "+", Literal(1))), _cmp(A, "<=", B)) is None


def test_float_offsets_do_not_round_into_contradictions() -> None:
    # 0.1 + 0.2 == 0.3 exactly in the reals.
    assert (
        _add(
            _cmp(A, ">=", Arith(B, "+", Literal(0.1))),
            _cmp(B, ">=", Arith(C, "+", Literal(0.2))),
            _cmp(C, ">=", Arith(A, "-", Literal(0.3))),
        )
        is not None
    )


def test_constant_only_comparison() -> None:
    assert _add(_cmp(Arith(A, "-", A), ">", Literal(0))) is None


def test_three_variable_terms_use_interval_bounds() -> None:
    fwd = _cmp(A, ">", Arith(B, "+", C))
    back = _cmp(B, ">", Arith(A, "+", C))
    nonneg = constraints_to_cell_type_env({"S!C1": Annotated[int, Between(0, 50)]}, {})
    anysign = constraints_to_cell_type_env({"S!C1": Annotated[float, RealBetween(-5, 5)]}, {})
    unbounded = constraints_to_cell_type_env({}, {})

    def feasible(env) -> bool:
        c = GuardConstraints().add(fwd, cell_type_env=env)
        return c is not None and c.add(back, cell_type_env=env) is not None

    assert feasible(nonneg) is False
    assert feasible(anysign) is True
    assert feasible(unbounded) is True


def test_interval_domains_bound_single_cells() -> None:
    env = constraints_to_cell_type_env({"S!A1": Annotated[int, Between(0, 10)]}, {})
    assert GuardConstraints().add(_cmp(A, ">", Literal(10)), cell_type_env=env) is None
    assert GuardConstraints().add(_cmp(A, ">=", Literal(10)), cell_type_env=env) is not None


def test_non_numeric_domain_is_not_ordered_numerically() -> None:
    """Text sorts above every number in Excel, so `A > B + 1` holds for any text A."""
    env = constraints_to_cell_type_env({"S!A1": TypingLiteral["x", "y"]}, {})
    c = GuardConstraints().add(_cmp(A, ">", Arith(B, "+", Literal(1))), cell_type_env=env)
    assert c is not None
    assert c.add(_cmp(B, ">", Arith(A, "+", Literal(1))), cell_type_env=env) is not None


def test_constraints_with_same_bounds_compare_equal() -> None:
    """DFS state dedup relies on equal constraint sets hashing equal."""
    one = _add(_cmp(A, ">", B), _cmp(B, ">", C))
    two = _add(_cmp(B, ">", C), _cmp(B, "<", A))
    assert one == two
    assert hash(one) == hash(two)


# --- Cycle reports ----------------------------------------------------------


def _schedule_cycle_graph() -> DependencyGraph:
    """Two payment cells, each live only when its year is after the other's plus grace.

    `P1` reads `P2` when `T1 > T2 + G`; `P2` reads `P1` when `T2 > T1 + G`.
    """
    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "P", 1, formula="=IF(T1>T2+G1,P2,0)", is_leaf=False))
    graph.add_node(make_cell_node("S", "P", 2, formula="=IF(T2>T1+G1,P1,0)", is_leaf=False))
    t1, t2, g = CellRef("S!T1"), CellRef("S!T2"), CellRef("S!G1")
    graph.add_edge("S!P1", "S!P2", guard=_cmp(t1, ">", Arith(t2, "+", g)))
    graph.add_edge("S!P2", "S!P1", guard=_cmp(t2, ">", Arith(t1, "+", g)))
    return graph


def test_schedule_pattern_is_acyclic_for_non_negative_grace() -> None:
    graph = _schedule_cycle_graph()
    env = constraints_to_cell_type_env({"S!G1": Annotated[int, Between(0, 50)]}, {})
    assert graph.cycle_report(cell_type_env=env).has_may_cycles is False


def test_schedule_pattern_may_cycle_when_grace_can_be_negative() -> None:
    graph = _schedule_cycle_graph()
    env = constraints_to_cell_type_env({"S!G1": Annotated[int, Between(-5, 50)]}, {})
    assert graph.cycle_report(cell_type_env=env).has_may_cycles is True
