"""Unit tests for constant-folding compression rule."""

from __future__ import annotations

from excel_grapher.compression.constant_folding import (
    apply_constant_folding,
    fold_literals_in_ast,
    try_fold_ast,
)
from excel_grapher.compression.parity import assert_compression_parity
from excel_grapher.compression.stats import empty_compression_stats
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    StringNode,
)
from excel_grapher.core.types import XlError

from .conftest import parse_formula


def test_try_fold_ast_simple_arithmetic() -> None:
    folded = try_fold_ast(parse_formula("=2+3"))
    assert folded == NumberNode(5.0)


def test_try_fold_ast_returns_none_for_refs() -> None:
    assert try_fold_ast(parse_formula("=Sheet1!A1+1")) is None


def test_try_fold_ast_returns_none_for_atomic_literals() -> None:
    assert try_fold_ast(NumberNode(1.0)) is None
    assert try_fold_ast(StringNode("x")) is None


def test_fold_literals_in_ast_nested_expression() -> None:
    ast = parse_formula("=(2+3)*4")
    assert fold_literals_in_ast(ast) == NumberNode(20.0)


def test_fold_literals_in_ast_partial_with_reference() -> None:
    ast = parse_formula("=2+3+Sheet1!A1")
    assert fold_literals_in_ast(ast) == BinaryOpNode(
        "+",
        NumberNode(5.0),
        CellRefNode("Sheet1!A1"),
    )


def test_fold_literals_in_ast_string_concatenation() -> None:
    ast = parse_formula('="Hello"&" "&"World"')
    assert fold_literals_in_ast(ast) == StringNode("Hello World")


def test_fold_literals_in_ast_division_by_zero() -> None:
    ast = parse_formula("=1/0")
    assert fold_literals_in_ast(ast) == ErrorNode(XlError.DIV)


def test_fold_literals_in_ast_leaves_unsupported_functions() -> None:
    ast = FunctionCallNode("FOO", [NumberNode(1.0), NumberNode(2.0)])
    assert fold_literals_in_ast(ast) == ast


def test_apply_constant_folding_records_stats() -> None:
    ast_map = {
        "Sheet1!A1": parse_formula("=2+3"),
        "Sheet1!B1": parse_formula("=Sheet1!C1+1"),
    }
    stats = empty_compression_stats()
    result = apply_constant_folding(ast_map, stats=stats)
    assert result["Sheet1!A1"] == NumberNode(5.0)
    assert result["Sheet1!B1"] == parse_formula("=Sheet1!C1+1")
    contrib = stats.contribution_for("constant_folding")
    assert contrib.in_place_transforms == 1
    assert contrib.cells_affected == 1


def test_apply_constant_folding_parity() -> None:
    input_values = {"Sheet1!B1": 4}
    original = {
        "Sheet1!A1": parse_formula("=2+3"),
        "Sheet1!C1": parse_formula("=5*Sheet1!B1"),
    }
    compressed = apply_constant_folding(original)
    assert_compression_parity(original, compressed, input_values=input_values)


def test_apply_constant_folding_parity_with_partial_fold() -> None:
    input_values = {"Sheet1!A1": 10}
    original = {"Sheet1!B1": parse_formula("=2+3+Sheet1!A1")}
    compressed = apply_constant_folding(original)
    assert_compression_parity(original, compressed, input_values=input_values)
