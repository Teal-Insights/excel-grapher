"""Rewrite formula ASTs by resolved cell identity (#544)."""

from __future__ import annotations

from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    NumberNode,
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
