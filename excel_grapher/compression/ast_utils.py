"""Shared AST walk helpers for compression rules."""

from __future__ import annotations

from collections.abc import Callable

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    ColumnVarCellRefNode,
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    SubexpressionRefNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
)


def map_ast(node: AstNode, transform: Callable[[AstNode], AstNode]) -> AstNode:
    """Return a copy of `node` with `transform` applied bottom-up."""
    replacement = transform(node)
    if replacement is not node:
        return replacement

    if isinstance(node, FunctionCallNode):
        return FunctionCallNode(node.name, [map_ast(arg, transform) for arg in node.args])
    if isinstance(node, BinaryOpNode):
        return BinaryOpNode(
            node.op,
            map_ast(node.left, transform),
            map_ast(node.right, transform),
        )
    if isinstance(node, UnaryOpNode):
        return UnaryOpNode(node.op, map_ast(node.operand, transform))
    return node


def is_literal_ast(node: AstNode) -> bool:
    """Return True when `node` contains no cell or range references."""
    if isinstance(node, (NumberNode, StringNode, BoolNode, ErrorNode, EmptyArgNode)):
        return True
    if isinstance(
        node,
        (
            CellRefNode,
            RangeNode,
            WholeColumnNode,
            WholeRowNode,
            ColumnVarCellRefNode,
            SubexpressionRefNode,
        ),
    ):
        return False
    if isinstance(node, FunctionCallNode):
        return all(is_literal_ast(arg) for arg in node.args)
    if isinstance(node, BinaryOpNode):
        return is_literal_ast(node.left) and is_literal_ast(node.right)
    if isinstance(node, UnaryOpNode):
        return is_literal_ast(node.operand)
    return False


def ast_contains_refs(node: AstNode) -> bool:
    """Return True when `node` contains any reference-like AST child."""
    if isinstance(
        node,
        (
            CellRefNode,
            RangeNode,
            WholeColumnNode,
            WholeRowNode,
            ColumnVarCellRefNode,
            SubexpressionRefNode,
        ),
    ):
        return True
    if isinstance(node, FunctionCallNode):
        return any(ast_contains_refs(arg) for arg in node.args)
    if isinstance(node, BinaryOpNode):
        return ast_contains_refs(node.left) or ast_contains_refs(node.right)
    if isinstance(node, UnaryOpNode):
        return ast_contains_refs(node.operand)
    return False
