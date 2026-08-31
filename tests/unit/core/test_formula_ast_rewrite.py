"""Rewrite formula ASTs by resolved cell identity (#544, #557)."""

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
