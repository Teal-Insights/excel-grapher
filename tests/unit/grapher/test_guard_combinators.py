"""Unit tests for the guard combinators `and_guard` and `or_guard`."""

from __future__ import annotations

from excel_grapher.grapher.guard import (
    And,
    CellRef,
    Compare,
    Literal,
    Not,
    Or,
    and_guard,
    or_guard,
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
