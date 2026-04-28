from __future__ import annotations

from excel_grapher.grapher.guard import CellRef, Compare, GuardConstraints, Literal, Not, Or


def test_guard_constraints_detects_contradiction_with_double_negation_equivalent() -> None:
    """
    Equivalent logical forms should lead to the same contradiction outcome.
    """
    key = "Sheet1!A1"
    neq_one = Not(Compare(left=CellRef(key=key), op="=", right=Literal(value=1)))
    eq_one_via_double_neg = Not(Not(Compare(left=CellRef(key=key), op="=", right=Literal(value=1))))

    c = GuardConstraints()
    c_after_neq = c.add(neq_one)
    assert c_after_neq is not None
    # Must be detected as contradictory once double negation is normalized away.
    assert c_after_neq.add(eq_one_via_double_neg) is None


def test_guard_constraints_normalizes_nested_and_or_variants() -> None:
    """
    Recursive normalization through AND/OR should preserve feasibility semantics.
    """
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
