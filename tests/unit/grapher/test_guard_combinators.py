"""Unit tests for the guard combinators `and_guard` and `or_guard`."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.guard import (
    And,
    CellRef,
    Compare,
    Literal,
    Not,
    Or,
    RangeRef,
    and_guard,
    intern_guard,
    or_guard,
    rewrite_guard_keys,
)

_A = Compare(left=CellRef(key="Sheet1!A1"), op="=", right=Literal(value=1))
_B = Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=2))
_C = Not(operand=Compare(left=CellRef(key="Sheet1!C1"), op="=", right=Literal(value=3)))


def test_and_guard_combines_two_simple_guards() -> None:
    assert and_guard(_A, _B) == And(operands=(_A, _B))


def test_and_guard_flattens_nested_ands_on_either_side() -> None:
    assert and_guard(And(operands=(_A, _B)), _C) == And(operands=(_A, _B, _C))
    assert and_guard(_A, And(operands=(_B, _C))) == And(operands=(_A, _B, _C))
    assert and_guard(And(operands=(_A,)), And(operands=(_B, _C))) == And(operands=(_A, _B, _C))


def test_and_guard_treats_literal_true_as_identity() -> None:
    assert and_guard(Literal(value=True), _A) == _A
    assert and_guard(_A, Literal(value=True)) == _A
    assert and_guard(Literal(value=True), Literal(value=True)) == Literal(value=True)


def test_and_guard_does_not_flatten_or_operands() -> None:
    disj = Or(operands=(_A, _B))
    assert and_guard(disj, _C) == And(operands=(disj, _C))


def test_or_guard_flattens_nested_ors() -> None:
    assert or_guard(Or(operands=(_A, _B)), _C) == Or(operands=(_A, _B, _C))


def test_rewrite_guard_keys_cell_ref_and_compare() -> None:
    expr = intern_guard(Compare(left=CellRef("Sheet1!A1"), op=">", right=Literal(0)))
    got = rewrite_guard_keys(expr, "Sheet1!A1", "Sheet1!D5")
    assert got == intern_guard(Compare(left=CellRef("Sheet1!D5"), op=">", right=Literal(0)))


def test_rewrite_guard_keys_unchanged_identity() -> None:
    expr = intern_guard(Compare(left=CellRef("Sheet1!A1"), op="=", right=Literal(1)))
    assert rewrite_guard_keys(expr, "Sheet1!Z9", "Sheet1!B2") is expr


def test_rewrite_guard_keys_walks_and_or_not() -> None:
    inner = intern_guard(Compare(left=CellRef("Sheet1!A1"), op=">", right=Literal(0)))
    expr = intern_guard(Not(inner))
    got = rewrite_guard_keys(expr, "Sheet1!A1", "Sheet1!D5")
    rewritten = intern_guard(Compare(left=CellRef("Sheet1!D5"), op=">", right=Literal(0)))
    assert got == intern_guard(Not(rewritten))


def test_rewrite_guard_keys_range_endpoints_not_interior() -> None:
    rng = intern_guard(RangeRef("Sheet1!A1:A3"))
    assert rewrite_guard_keys(rng, "Sheet1!A2", "Sheet1!C9") is rng
    assert rewrite_guard_keys(rng, "Sheet1!A1", "Sheet1!D5") == intern_guard(
        RangeRef("Sheet1!D5:A3")
    )
    assert rewrite_guard_keys(rng, "Sheet1!A3", "Sheet1!D5") == intern_guard(
        RangeRef("Sheet1!A1:D5")
    )


def test_rewrite_guard_keys_range_rejects_cross_sheet() -> None:
    rng = intern_guard(RangeRef("Sheet1!A1:A3"))
    with pytest.raises(ValueError, match="sheet"):
        rewrite_guard_keys(rng, "Sheet1!A1", "Sheet2!B1")
