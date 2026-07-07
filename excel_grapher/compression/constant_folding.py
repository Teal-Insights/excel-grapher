"""Rule 3: constant folding for literal-only subexpressions."""

from __future__ import annotations

from collections.abc import Mapping

from excel_grapher.core.address_keys import normalize_key
from excel_grapher.core.expr_eval import Unsupported, evaluate_expr
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    StringNode,
    UnaryOpNode,
)
from excel_grapher.core.types import CellValue, XlError

from .ast_utils import is_literal_ast
from .stats import CompressionStats


def try_fold_ast(ast: AstNode) -> AstNode | None:
    """Fold a literal-only AST to a single literal node, if possible."""
    if not is_literal_ast(ast):
        return None
    if isinstance(ast, (NumberNode, StringNode, BoolNode, ErrorNode)):
        return None

    result = evaluate_expr(ast, get_cell_value=_literal_fold_cell_ref)
    if isinstance(result, Unsupported):
        return None
    return _value_to_ast(result)


def fold_literals_in_ast(ast: AstNode) -> AstNode:
    """Fold literal subexpressions until the AST reaches a fixed point."""
    while True:
        folded = _fold_literals_once(ast)
        if folded == ast:
            return ast
        ast = folded


def apply_constant_folding(
    ast_map: Mapping[str, AstNode],
    stats: CompressionStats | None = None,
) -> dict[str, AstNode]:
    """Apply constant folding to every formula in `ast_map`."""
    result: dict[str, AstNode] = {}
    transforms = 0

    for cell_key, ast in ast_map.items():
        folded = fold_literals_in_ast(ast)
        if folded != ast:
            transforms += 1
        result[normalize_key(cell_key)] = folded

    if stats is not None:
        stats.contribution_for("constant_folding").record(
            in_place_transforms=transforms,
            cells_affected=transforms,
        )
    return result


def _fold_literals_once(ast: AstNode) -> AstNode:
    if isinstance(ast, FunctionCallNode):
        ast = FunctionCallNode(
            ast.name,
            [_fold_literals_once(arg) for arg in ast.args],
        )
    elif isinstance(ast, BinaryOpNode):
        ast = BinaryOpNode(
            ast.op,
            _fold_literals_once(ast.left),
            _fold_literals_once(ast.right),
        )
    elif isinstance(ast, UnaryOpNode):
        ast = UnaryOpNode(ast.op, _fold_literals_once(ast.operand))

    replacement = try_fold_ast(ast)
    return replacement if replacement is not None else ast


def _literal_fold_cell_ref(_address: str) -> CellValue:
    raise AssertionError("literal fold should not resolve cell references")


def _value_to_ast(value: CellValue | XlError) -> AstNode:
    if isinstance(value, XlError):
        return ErrorNode(value)
    if isinstance(value, bool):
        return BoolNode(value)
    if isinstance(value, str):
        return StringNode(value)
    if isinstance(value, (int, float)):
        return NumberNode(float(value))
    raise TypeError(f"Unsupported folded value type: {type(value)!r}")
