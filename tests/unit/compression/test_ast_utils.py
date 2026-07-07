"""Unit tests for compression AST utilities."""

from __future__ import annotations

from excel_grapher.compression.ast_utils import ast_contains_refs, is_literal_ast, map_ast
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    ColumnVarCellRefNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    SubexpressionRefNode,
    UnaryOpNode,
)
from excel_grapher.core.types import XlError

from .conftest import parse_formula


def test_map_ast_identity_preserves_tree() -> None:
    ast = parse_formula("=Sheet1!A1+Sheet1!B1*2")
    assert map_ast(ast, lambda node: node) == ast


def test_map_ast_replaces_matching_nodes() -> None:
    ast = BinaryOpNode("+", CellRefNode("Sheet1!A1"), NumberNode(1.0))
    mapped = map_ast(
        ast,
        lambda node: CellRefNode("Sheet1!B1") if isinstance(node, CellRefNode) else node,
    )
    assert mapped == BinaryOpNode("+", CellRefNode("Sheet1!B1"), NumberNode(1.0))


def test_map_ast_descends_into_nested_calls() -> None:
    inner = CellRefNode("Sheet1!A1")
    ast = FunctionCallNode("IF", [BoolNode(True), inner, NumberNode(2.0)])

    def _double_numbers(node):
        if isinstance(node, NumberNode):
            return NumberNode(node.value * 2)
        return node

    mapped = map_ast(ast, _double_numbers)
    assert mapped == FunctionCallNode("IF", [BoolNode(True), inner, NumberNode(4.0)])


def test_map_ast_handles_unary_operators() -> None:
    ast = UnaryOpNode("+", CellRefNode("Sheet1!A1"))
    mapped = map_ast(
        ast,
        lambda node: NumberNode(5.0) if isinstance(node, CellRefNode) else node,
    )
    assert mapped == UnaryOpNode("+", NumberNode(5.0))


def test_is_literal_ast_for_atomic_literals() -> None:
    assert is_literal_ast(NumberNode(1.0))
    assert is_literal_ast(StringNode("x"))
    assert is_literal_ast(BoolNode(True))
    assert is_literal_ast(ErrorNode(XlError.NA))


def test_is_literal_ast_for_literal_expressions() -> None:
    assert is_literal_ast(parse_formula("=2+3"))
    assert is_literal_ast(parse_formula('="Hello"&" World"'))
    assert is_literal_ast(
        FunctionCallNode("IF", [BoolNode(True), NumberNode(1.0), NumberNode(2.0)])
    )


def test_is_literal_ast_rejects_references() -> None:
    assert not is_literal_ast(CellRefNode("Sheet1!A1"))
    assert not is_literal_ast(RangeNode("Sheet1!A1", "Sheet1!B2"))
    assert not is_literal_ast(ColumnVarCellRefNode(sheet="Ext", row=87))
    assert not is_literal_ast(SubexpressionRefNode("_cse!0"))
    assert not is_literal_ast(BinaryOpNode("+", NumberNode(2.0), CellRefNode("Sheet1!A1")))


def test_ast_contains_refs_detects_reference_nodes() -> None:
    assert not ast_contains_refs(NumberNode(1.0))
    assert ast_contains_refs(CellRefNode("Sheet1!A1"))
    assert ast_contains_refs(RangeNode("Sheet1!A1", "Sheet1!B2"))
    assert ast_contains_refs(ColumnVarCellRefNode())
    assert ast_contains_refs(SubexpressionRefNode("_cse!0"))


def test_ast_contains_refs_detects_nested_references() -> None:
    literal_if = FunctionCallNode("IF", [BoolNode(True), NumberNode(1.0), NumberNode(2.0)])
    ref_if = FunctionCallNode(
        "IF",
        [BoolNode(True), NumberNode(1.0), CellRefNode("Sheet1!A1")],
    )
    assert not ast_contains_refs(literal_if)
    assert ast_contains_refs(ref_if)
