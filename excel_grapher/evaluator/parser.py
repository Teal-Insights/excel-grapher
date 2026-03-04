from __future__ import annotations

# ruff: noqa: I001

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    ErrorNode,
    FormulaParseError,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
    parse as _core_parse,
)

from .errors import ParseError


__all__ = [
    "AstNode",
    "BinaryOpNode",
    "BoolNode",
    "CellRefNode",
    "ErrorNode",
    "FunctionCallNode",
    "NumberNode",
    "RangeNode",
    "StringNode",
    "UnaryOpNode",
    "parse",
]


def parse(formula: str) -> AstNode:
    """Wrapper around the core formula parser that preserves evaluator.ParseError."""

    try:
        return _core_parse(formula)
    except FormulaParseError as exc:
        raise ParseError(exc.formula, exc.message) from exc
