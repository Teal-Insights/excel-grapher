"""Rewrite formula ASTs by resolved cell identity (#544, #549, #557)."""

from __future__ import annotations

from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    BinaryOpNode,
    CellRef,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    RelativeAxis,
    WholeColumnNode,
    WholeRowNode,
    ast_mentions_resolved_non_cell_key,
    parse,
    parse_preserving_axes,
    rebase_relative_axes,
    replace_resolved_cell_ref,
    resolve_cell_ref,
    resolve_whole_column_ref,
    resolve_whole_row_ref,
    retarget_resolved_refs,
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


def test_replace_resolved_cell_ref_retargets_range_endpoints() -> None:
    ast = parse("=SUM(Sheet1!B1:B3)+Sheet1!B1")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C1", anchor="Sheet1!A1"
    )
    assert rewritten == parse("=SUM(Sheet1!C1:B3)+Sheet1!C1")
    assert unparse_normalized_formula(rewritten) == "=SUM(Sheet1!C1:B3)+Sheet1!C1"


def test_replace_resolved_cell_ref_retargets_range_end() -> None:
    ast = parse("=SUM(Sheet1!A1:B1)")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!D4", anchor="Sheet1!A1"
    )
    assert rewritten == parse("=SUM(Sheet1!A1:D4)")


def test_replace_resolved_cell_ref_retargets_cross_sheet_range_start() -> None:
    ast = parse("=Sheet1!B1:Sheet2!B10")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C1", anchor="Sheet1!A1"
    )
    assert rewritten == parse("=Sheet1!C1:Sheet2!B10")


def test_replace_resolved_cell_ref_preserves_range_endpoint_axes() -> None:
    ast = parse_preserving_axes("=SUM(B1:B3)", anchor="Sheet1!A1")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C1", anchor="Sheet1!A1"
    )
    assert isinstance(rewritten, FunctionCallNode)
    rng = rewritten.args[0]
    assert isinstance(rng, RangeNode)
    assert rng.start_ref == CellRef(sheet="Sheet1", col=RelativeAxis(2), row=RelativeAxis(0))
    assert rng.end_ref == CellRef(sheet="Sheet1", col=RelativeAxis(1), row=RelativeAxis(2))
    assert unparse_normalized_formula(rewritten, anchor="Sheet1!A1") == "=SUM(Sheet1!C1:B3)"


def test_replace_resolved_cell_ref_retargets_whole_column() -> None:
    ast = parse("=SUM(Sheet1!B:B)")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C2", anchor="Sheet1!A1"
    )
    assert rewritten == parse("=SUM(Sheet1!C:C)")
    assert isinstance(rewritten, FunctionCallNode)
    col = rewritten.args[0]
    assert isinstance(col, WholeColumnNode)
    assert col == WholeColumnNode(sheet="Sheet1", column="C")


def test_replace_resolved_cell_ref_preserves_whole_column_relative_axis() -> None:
    ast = parse_preserving_axes("=SUM(B:B)", anchor="Sheet1!A1")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C1", anchor="Sheet1!A1"
    )
    assert isinstance(rewritten, FunctionCallNode)
    col = rewritten.args[0]
    assert isinstance(col, WholeColumnNode)
    assert col == WholeColumnNode(sheet="Sheet1", col=RelativeAxis(2))
    assert unparse_normalized_formula(rewritten, anchor="Sheet1!A1") == "=SUM(Sheet1!C:C)"


def test_replace_resolved_cell_ref_retargets_whole_row() -> None:
    ast = parse("=SUM(Sheet1!1:1)")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C4", anchor="Sheet1!A2"
    )
    assert rewritten == parse("=SUM(Sheet1!4:4)")
    assert isinstance(rewritten, FunctionCallNode)
    row = rewritten.args[0]
    assert isinstance(row, WholeRowNode)
    assert row == WholeRowNode(sheet="Sheet1", row=4)


def test_replace_resolved_cell_ref_preserves_whole_row_relative_axis() -> None:
    ast = parse_preserving_axes("=SUM(1:1)", anchor="Sheet1!A2")
    rewritten = replace_resolved_cell_ref(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C4", anchor="Sheet1!A2"
    )
    assert isinstance(rewritten, FunctionCallNode)
    row = rewritten.args[0]
    assert isinstance(row, WholeRowNode)
    assert row == WholeRowNode(sheet="Sheet1", row=RelativeAxis(2))
    assert unparse_normalized_formula(rewritten, anchor="Sheet1!A2") == "=SUM(Sheet1!4:4)"


def test_replace_does_not_splice_subtree_into_range_or_whole_leaves() -> None:
    host = parse("=SUM(Sheet1!B1:B3)+Sheet1!B1")
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
        FunctionCallNode("SUM", [parse("=Sheet1!B1:B3")]),
        BinaryOpNode("*", CellRefNode("Sheet1!D1"), NumberNode(2.0)),
    )


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


def test_rebase_relative_axes_keeps_resolved_targets() -> None:
    ast = parse_preserving_axes("=A1+1", anchor="Sheet1!B2")
    rebased = rebase_relative_axes(ast, old_anchor="Sheet1!B2", new_anchor="Sheet1!C3")
    assert unparse_normalized_formula(rebased, anchor="Sheet1!C3") == "=Sheet1!A1+1"
    assert isinstance(rebased, BinaryOpNode)
    left = rebased.left
    assert isinstance(left, CellRefNode)
    assert left.ref.col == RelativeAxis(-2)
    assert left.ref.row == RelativeAxis(-2)
    assert resolve_cell_ref(left.ref, "Sheet1!C3") == "Sheet1!A1"


def test_rebase_relative_axes_mixed_and_absolute() -> None:
    ast = parse_preserving_axes("=$A1+A$1+$B$2", anchor="Sheet1!B2")
    rebased = rebase_relative_axes(ast, old_anchor="Sheet1!B2", new_anchor="Sheet1!C4")
    assert (
        unparse_normalized_formula(rebased, anchor="Sheet1!C4") == "=Sheet1!A1+Sheet1!A1+Sheet1!B2"
    )
    assert isinstance(rebased, BinaryOpNode)
    mixed_col = rebased.left
    assert isinstance(mixed_col, BinaryOpNode)
    dollar_col = mixed_col.left
    dollar_row = mixed_col.right
    assert isinstance(dollar_col, CellRefNode)
    assert dollar_col.ref.col == AbsoluteAxis(1)
    assert dollar_col.ref.row == RelativeAxis(-3)
    assert isinstance(dollar_row, CellRefNode)
    assert dollar_row.ref.col == RelativeAxis(-2)
    assert dollar_row.ref.row == AbsoluteAxis(1)
    assert isinstance(rebased.right, CellRefNode)
    assert rebased.right.ref.col == AbsoluteAxis(2)
    assert rebased.right.ref.row == AbsoluteAxis(2)


def test_rebase_relative_axes_range_and_whole_leaves() -> None:
    rng = parse_preserving_axes("=SUM(A1:A3)", anchor="Sheet1!B4")
    rebased_rng = rebase_relative_axes(rng, old_anchor="Sheet1!B4", new_anchor="Sheet1!D6")
    assert unparse_normalized_formula(rebased_rng, anchor="Sheet1!D6") == "=SUM(Sheet1!A1:A3)"
    assert isinstance(rebased_rng, FunctionCallNode)
    inner = rebased_rng.args[0]
    assert isinstance(inner, RangeNode)
    assert inner.start_ref.col == RelativeAxis(-3)
    assert inner.start_ref.row == RelativeAxis(-5)
    assert inner.end_ref.col == RelativeAxis(-3)
    assert inner.end_ref.row == RelativeAxis(-3)

    col = parse_preserving_axes("=A:A", anchor="Sheet1!B2")
    rebased_col = rebase_relative_axes(col, old_anchor="Sheet1!B2", new_anchor="Sheet1!C3")
    assert isinstance(rebased_col, WholeColumnNode)
    assert rebased_col.col == RelativeAxis(-2)
    assert resolve_whole_column_ref(rebased_col, "Sheet1!C3") == ("Sheet1", "A")

    row = parse_preserving_axes("=1:1", anchor="Sheet1!A3")
    rebased_row = rebase_relative_axes(row, old_anchor="Sheet1!A3", new_anchor="Sheet1!B5")
    assert isinstance(rebased_row, WholeRowNode)
    assert rebased_row.row == RelativeAxis(-4)
    assert resolve_whole_row_ref(rebased_row, "Sheet1!B5") == ("Sheet1", 1)


def test_rebase_relative_axes_noop_when_anchor_unchanged() -> None:
    ast = parse_preserving_axes("=A1", anchor="Sheet1!B2")
    assert rebase_relative_axes(ast, old_anchor="Sheet1!B2", new_anchor="Sheet1!B2") is ast


def test_retarget_resolved_refs_preserves_identity_for_interior_occupancy() -> None:
    ast = parse_preserving_axes("=SUM(A1:A3)", anchor="Sheet1!B1")
    rewritten = retarget_resolved_refs(
        ast, old_key="Sheet1!A2", new_key="Sheet1!C9", anchor="Sheet1!B1"
    )
    assert rewritten is ast


def test_rebase_relative_axes_preserves_function_identity_without_relative_refs() -> None:
    ast = parse_preserving_axes("=SUM(1,2)", anchor="Sheet1!B2")
    rebased = rebase_relative_axes(ast, old_anchor="Sheet1!B2", new_anchor="Sheet1!C3")
    assert rebased is ast


def test_retarget_resolved_refs_range_endpoints() -> None:
    ast = parse_preserving_axes("=SUM(B1:B3)+B1", anchor="Sheet1!A1")
    rewritten = retarget_resolved_refs(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C2", anchor="Sheet1!A1"
    )
    assert unparse_normalized_formula(rewritten, anchor="Sheet1!A1") == (
        "=SUM(Sheet1!C2:B3)+Sheet1!C2"
    )
    assert isinstance(rewritten, BinaryOpNode)
    assert isinstance(rewritten.left, FunctionCallNode)
    rng = rewritten.left.args[0]
    assert isinstance(rng, RangeNode)
    assert rng.start_ref.col == RelativeAxis(2)
    assert rng.start_ref.row == RelativeAxis(1)
    assert rng.end_ref.col == RelativeAxis(1)
    assert rng.end_ref.row == RelativeAxis(2)


def test_retarget_resolved_refs_mixed_axes() -> None:
    ast = CellRefNode(CellRef(sheet="Sheet1", col=AbsoluteAxis(2), row=RelativeAxis(0)))
    rewritten = retarget_resolved_refs(
        ast, old_key="Sheet1!B1", new_key="Sheet1!C2", anchor="Sheet1!A1"
    )
    assert isinstance(rewritten, CellRefNode)
    assert rewritten.ref == CellRef(sheet="Sheet1", col=AbsoluteAxis(3), row=RelativeAxis(1))


def test_ast_mentions_resolved_non_cell_key_range_and_whole() -> None:
    mixed = parse("=SUM(Sheet1!B1:B3)+Sheet1!B1")
    assert ast_mentions_resolved_non_cell_key(mixed, key="Sheet1!B1", anchor="Sheet1!A1")
    assert not ast_mentions_resolved_non_cell_key(mixed, key="Sheet1!C1", anchor="Sheet1!A1")
    assert ast_mentions_resolved_non_cell_key(
        parse("=SUM(Sheet1!B:B)"), key="Sheet1!B5", anchor="Sheet1!A1"
    )
    assert ast_mentions_resolved_non_cell_key(
        parse("=SUM(Sheet1!1:1)"), key="Sheet1!Z1", anchor="Sheet1!A2"
    )
    assert not ast_mentions_resolved_non_cell_key(
        parse("=Sheet1!B1+1"), key="Sheet1!B1", anchor="Sheet1!A1"
    )
