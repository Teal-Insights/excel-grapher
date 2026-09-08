from __future__ import annotations

import random
from typing import Annotated
from typing import Literal as TypingLiteral

from excel_grapher.core.cell_types import Between, constraints_to_cell_type_env
from excel_grapher.grapher.guard import (
    CellRef,
    Compare,
    GuardConstraints,
    Literal,
    Not,
    Or,
    canonicalize_guard,
)


def _nest_not(expr: Compare, count: int) -> Compare | Not:
    out: Compare | Not = expr
    for _ in range(count):
        out = Not(out)
    return out


def test_guard_constraints_detects_contradiction_with_double_negation_equivalent() -> None:
    """Equivalent logical forms should lead to the same contradiction outcome."""
    key = "Sheet1!A1"
    neq_one = Not(Compare(left=CellRef(key=key), op="=", right=Literal(value=1)))
    eq_one_via_double_neg = Not(Not(Compare(left=CellRef(key=key), op="=", right=Literal(value=1))))

    c = GuardConstraints()
    c_after_neq = c.add(neq_one)
    assert c_after_neq is not None
    # Must be detected as contradictory once double negation is normalized away.
    assert c_after_neq.add(eq_one_via_double_neg) is None


def test_guard_constraints_normalizes_nested_and_or_variants() -> None:
    """Recursive normalization through AND/OR should preserve feasibility semantics."""
    key = "Sheet1!A1"
    eq_one = Compare(left=CellRef(key=key), op="=", right=Literal(value=1))
    neq_one_via_nested_not = Not(Not(Not(eq_one)))
    disj = Or(
        (
            Not(Not(eq_one)),
            neq_one_via_nested_not,
        )
    )

    c = GuardConstraints()
    c2 = c.add(disj)
    # OR remains opaque, but canonicalization should remove any double negations.
    assert c2 is not None
    assert len(c2.opaque) == 1
    assert "NOT(NOT(" not in c2.opaque[0]


def test_canonicalize_guard_random_not_depth_matches_odd_even_parity() -> None:
    key = "Sheet1!A1"
    eq_one = Compare(left=CellRef(key=key), op="=", right=Literal(value=1))
    rng = random.Random(20260427)
    not_depth = rng.randint(3, 10)

    normalized = canonicalize_guard(_nest_not(eq_one, not_depth))
    if not_depth % 2 == 0:
        assert normalized == eq_one
    else:
        assert normalized == Not(eq_one)


def test_guard_constraints_cell_cell_equality_contradicts_its_negation() -> None:
    """Compare(CellRef, CellRef) must participate in consistency, not stay opaque."""
    eq = Compare(left=CellRef(key="Sheet1!A1"), op="=", right=CellRef(key="Sheet1!B1"))
    seed = GuardConstraints().add(eq)
    assert seed is not None
    assert seed.add(Not(eq)) is None


def test_guard_constraints_accepts_literal_on_either_side() -> None:
    left_lit = Compare(left=Literal(value=0), op="=", right=CellRef(key="Sheet1!A1"))
    seed = GuardConstraints().add(left_lit)
    assert seed is not None
    assert seed.add(Compare(left=CellRef(key="Sheet1!A1"), op="=", right=Literal(value=1))) is None


def test_guard_constraints_enum_domain_makes_complementary_inequalities_unsat() -> None:
    env = constraints_to_cell_type_env({"Sheet1!B1": TypingLiteral[0, 1]}, {})
    ne0 = Not(Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=0)))
    ne1 = Not(Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=1)))
    seed = GuardConstraints().add(ne0, cell_type_env=env)
    assert seed is not None
    assert seed.add(ne1, cell_type_env=env) is None


def test_guard_constraints_rejects_equality_outside_interval() -> None:
    env = constraints_to_cell_type_env(
        {"Sheet1!C1": Annotated[int, Between(min=0, max=1)]},
        {},
    )
    eq2 = Compare(left=CellRef(key="Sheet1!C1"), op="=", right=Literal(value=2))
    assert GuardConstraints().add(eq2, cell_type_env=env) is None


def test_guard_constraints_quoted_guard_key_matches_unquoted_env_key() -> None:
    env = constraints_to_cell_type_env(
        {"'Input 5 - Local-debt Financing'!C78": TypingLiteral[0, 1]},
        {},
    )
    eq2 = Compare(
        left=CellRef(key="'Input 5 - Local-debt Financing'!C78"),
        op="=",
        right=Literal(value=2),
    )
    assert GuardConstraints().add(eq2, cell_type_env=env) is None
