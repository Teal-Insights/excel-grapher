"""Unparse formula ASTs to absolute A1 `normalized_formula` text (#544)."""

from __future__ import annotations

from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    NumberNode,
    parse,
    parse_preserving_axes,
    unparse_normalized_formula,
)


def test_unparse_literals_and_ops() -> None:
    assert unparse_normalized_formula(parse("=1+2")) == "=1+2"
    assert unparse_normalized_formula(parse("=1.5")) == "=1.5"
    assert unparse_normalized_formula(parse('="a""b"')) == '="a""b"'
    assert unparse_normalized_formula(parse("=TRUE")) == "=TRUE"
    assert unparse_normalized_formula(parse("=#N/A")) == "=#N/A"


def test_unparse_cell_and_range() -> None:
    assert unparse_normalized_formula(parse("=Sheet1!A1")) == "=Sheet1!A1"
    assert unparse_normalized_formula(parse("=SUM(Sheet1!A1:A3)")) == "=SUM(Sheet1!A1:A3)"
    assert unparse_normalized_formula(parse("='Other Sheet'!B5+1")) == "='Other Sheet'!B5+1"


def test_unparse_resolves_relative_axes_against_anchor() -> None:
    ast = parse_preserving_axes("=A1*2", anchor="Sheet1!B1")
    assert unparse_normalized_formula(ast, anchor="Sheet1!B1") == "=Sheet1!A1*2"
    assert unparse_normalized_formula(ast, anchor="Sheet1!B2") == "=Sheet1!A2*2"


def test_unparse_parenthesizes_for_precedence() -> None:
    ast = BinaryOpNode(
        "*",
        BinaryOpNode("+", CellRefNode("Sheet1!D1"), NumberNode(1.0)),
        NumberNode(2.0),
    )
    assert unparse_normalized_formula(ast) == "=(Sheet1!D1+1)*2"


def test_unparse_omits_redundant_parens_when_inner_binds_tighter() -> None:
    ast = BinaryOpNode(
        "+",
        BinaryOpNode("*", CellRefNode("Sheet1!D1"), NumberNode(2.0)),
        NumberNode(1.0),
    )
    assert unparse_normalized_formula(ast) == "=Sheet1!D1*2+1"


def test_unparse_function_empty_args() -> None:
    ast = parse("=INDEX(Sheet1!A1:B2,,1)")
    assert unparse_normalized_formula(ast) == "=INDEX(Sheet1!A1:B2,,1)"


def test_unparse_whole_column() -> None:
    ast = parse('=MATCH("x",Data!A:A,0)')
    assert unparse_normalized_formula(ast) == '=MATCH("x",Data!A:A,0)'


def test_unparse_unary_minus_and_percent() -> None:
    assert unparse_normalized_formula(parse("=-Sheet1!A1")) == "=-Sheet1!A1"
    assert unparse_normalized_formula(parse("=100%")) == "=100%"


def test_unparse_round_trip_parse() -> None:
    formulas = [
        "=Sheet1!A1+Sheet1!A2",
        "=IF(Sheet1!A1,Sheet1!B1,Sheet1!C1)",
        "=SUM(Sheet1!A1:B2)*2",
        "=(Sheet1!A1+Sheet1!A2)*Sheet1!A3",
    ]
    for formula in formulas:
        ast = parse(formula)
        rendered = unparse_normalized_formula(ast)
        assert parse(rendered) == ast
        assert unparse_normalized_formula(parse(rendered)) == rendered
