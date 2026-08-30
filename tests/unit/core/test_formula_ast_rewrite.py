"""Rewrite formula ASTs by resolved cell identity (#544)."""

from __future__ import annotations

from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    BinaryOpNode,
    CellRef,
    CellRefNode,
    NumberNode,
    RelativeAxis,
    parse,
    parse_preserving_axes,
    replace_resolved_cell_ref,
    unparse_normalized_formula,
)


def test_replace_resolved_cell_ref_absolute() -> None:
    ast = parse("=Sheet1!B1+1")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C1", anchor="Sheet1!A1"
    )
    assert rewritten == parse("=Sheet1!C1+1")
    assert unparse_normalized_formula(rewritten) == "=Sheet1!C1+1"


def test_replace_resolved_cell_ref_relative() -> None:
    ast = parse_preserving_axes("=B1+1", anchor="Sheet1!A1")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C1", anchor="Sheet1!A1"
    )
    assert unparse_normalized_formula(rewritten, anchor="Sheet1!A1") == "=Sheet1!C1+1"
    assert isinstance(rewritten, BinaryOpNode)
    left = rewritten.left
    assert isinstance(left, CellRefNode)
    assert isinstance(left.ref.col, RelativeAxis)
    assert left.ref.col.offset == 2
    assert isinstance(left.ref.row, RelativeAxis)
    assert left.ref.row.offset == 0


def test_replace_resolved_cell_ref_preserves_mixed_axes() -> None:
    ast = CellRefNode(CellRef(sheet="Sheet1", col=AbsoluteAxis(2), row=RelativeAxis(0)))
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C2", anchor="Sheet1!A1"
    )
    assert isinstance(rewritten, CellRefNode)
    assert rewritten.ref == CellRef(sheet="Sheet1", col=AbsoluteAxis(3), row=RelativeAxis(1))
    assert unparse_normalized_formula(rewritten, anchor="Sheet1!A1") == "=Sheet1!C2"


def test_replace_does_not_touch_range_endpoints_as_identity_sites() -> None:
    ast = parse("=SUM(Sheet1!B1:B3)+Sheet1!B1")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C1", anchor="Sheet1!A1"
    )
    assert rewritten == parse("=SUM(Sheet1!B1:B3)+Sheet1!C1")


def test_inline_replaces_cell_with_subtree() -> None:
    host = parse("=Sheet1!B1+1")
    body = parse("=Sheet1!D1*2")
    rewritten = replace_resolved_cell_ref(
        host,
        old_key="Sheet1!B1",
        new_key="Sheet1!C1",
        anchor="Sheet1!A1",
        replacement=body,
    )
    assert rewritten == BinaryOpNode(
        "+",
        BinaryOpNode("*", CellRefNode("Sheet1!D1"), NumberNode(2.0)),
        NumberNode(1.0),
    )
