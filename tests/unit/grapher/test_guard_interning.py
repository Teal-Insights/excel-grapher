"""Hash-consing / interning for GuardExpr trees (#491)."""

from __future__ import annotations

import gc

from excel_grapher.grapher.guard import (
    And,
    CellRef,
    Compare,
    Literal,
    Not,
    Or,
    RangeRef,
    and_guard,
    clear_guard_intern_pool,
    guard_intern_pool_size,
    intern_guard,
    or_guard,
)
from excel_grapher.grapher.parser import parse_guard_expr


def test_intern_guard_keeps_bool_and_int_literals_distinct() -> None:
    assert intern_guard(Literal(True)) is not intern_guard(Literal(1))
    assert intern_guard(Literal(False)) is not intern_guard(Literal(0))


def test_intern_guard_returns_identical_object_for_equal_trees() -> None:
    a = Compare(CellRef("Sheet1!A1"), "=", Literal(1))
    b = Compare(CellRef("Sheet1!A1"), "=", Literal(1))
    assert a is not b
    assert intern_guard(a) is intern_guard(b)


def test_intern_guard_shares_subexpressions() -> None:
    left = CellRef("Sheet1!A1")
    shared = intern_guard(Compare(left, "=", Literal(1)))
    wrapped = intern_guard(Not(Compare(CellRef("Sheet1!A1"), "=", Literal(1))))
    assert isinstance(wrapped, Not)
    assert wrapped.operand is shared


def test_guard_ast_nodes_are_slotted() -> None:
    for cls in (CellRef, RangeRef, Literal, Compare, Not, And, Or):
        assert hasattr(cls, "__slots__")
        instance = intern_guard(
            {
                CellRef: CellRef("Sheet1!A1"),
                RangeRef: RangeRef("Sheet1!A1:A2"),
                Literal: Literal(0),
                Compare: Compare(CellRef("Sheet1!A1"), ">", Literal(0)),
                Not: Not(Literal(True)),
                And: And((Literal(True), Literal(False))),
                Or: Or((Literal(True), Literal(False))),
            }[cls]
        )
        assert not hasattr(instance, "__dict__")


def test_and_or_combinators_return_interned_results() -> None:
    a = Compare(CellRef("Sheet1!A1"), "=", Literal(1))
    b = Compare(CellRef("Sheet1!B1"), "=", Literal(2))
    first = and_guard(a, b)
    second = and_guard(
        Compare(CellRef("Sheet1!A1"), "=", Literal(1)),
        Compare(CellRef("Sheet1!B1"), "=", Literal(2)),
    )
    assert first is second
    assert or_guard(a, b) is or_guard(
        Compare(CellRef("Sheet1!A1"), "=", Literal(1)),
        Compare(CellRef("Sheet1!B1"), "=", Literal(2)),
    )


def test_parse_guard_expr_interns_identical_conditions() -> None:
    first = parse_guard_expr("$A$1=1", current_sheet="Sheet1")
    second = parse_guard_expr("$A$1=1", current_sheet="Sheet1")
    assert first is not None and second is not None
    assert first is second


def test_unreferenced_interned_guards_can_be_collected() -> None:
    """Pool entries are weak: dropping all strong refs frees the interned tree."""
    marker = "unique-gc-marker-zz999"
    expr = intern_guard(Compare(CellRef("Sheet1!ZZ999"), "=", Literal(marker)))
    expr_id = id(expr)
    del expr
    gc.collect()
    revived = intern_guard(Compare(CellRef("Sheet1!ZZ999"), "=", Literal(marker)))
    # Old pooled object was collected; this is a new interned instance.
    assert id(revived) != expr_id
    twin = intern_guard(Compare(CellRef("Sheet1!ZZ999"), "=", Literal(marker)))
    assert twin is revived


def test_clear_guard_intern_pool_drops_entries_but_reseeds_true() -> None:
    held = intern_guard(Compare(CellRef("Sheet1!AA1"), "=", Literal(42)))
    clear_guard_intern_pool()
    # Cleared entries are gone; a new equal tree is a distinct object from `held`.
    revived = intern_guard(Compare(CellRef("Sheet1!AA1"), "=", Literal(42)))
    assert revived == held
    assert revived is not held
    # and_guard identity for TRUE still works after reseed.
    assert and_guard(Literal(True), revived) is revived
    assert guard_intern_pool_size() >= 1
