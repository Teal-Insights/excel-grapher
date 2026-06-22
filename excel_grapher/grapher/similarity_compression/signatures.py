"""Structural signatures for clustering collapsed formula roots."""

from __future__ import annotations

import re
from dataclasses import dataclass

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FormulaParseError,
    FunctionCallNode,
    RangeNode,
    UnaryOpNode,
    parse,
)

from ..graph import DependencyGraph, NodeKey

__all__ = [
    "StructuralSignature",
    "classify_formula_kind",
    "function_skeleton",
    "structural_signature",
]


@dataclass(frozen=True)
class StructuralSignature:
    """Coarse bucket key for parallel-family pre-clustering."""

    kind: str
    row_band: int
    dependency_count: int
    function_skeleton: str


def _root_row_column(root: NodeKey) -> tuple[int, str]:
    _, rest = root.split("!", 1)
    if rest.startswith("'"):
        rest = rest.split("!", 1)[-1]
    column = "".join(char for char in rest if char.isalpha())
    row = int("".join(char for char in rest if char.isdigit()))
    return row, column.upper()


def classify_formula_kind(formula: str) -> str:
    """Return a coarse formula-family tag from normalized text."""
    upper = formula.upper()
    if "CHOOSE" in upper and "/100" in upper and "*" in upper:
        return "shocked_year_block"
    if "CHOOSE" not in upper and "/100" not in upper and re.search(r"[+-]", upper):
        return "linear_aggregate"
    return "generic"


def _dependency_count(ast: AstNode) -> int:
    count = 0

    def walk(node: AstNode) -> None:
        nonlocal count
        if isinstance(node, CellRefNode):
            count += 1
        elif isinstance(node, RangeNode):
            count += 2
        elif isinstance(node, FunctionCallNode):
            for arg in node.args:
                walk(arg)
        elif isinstance(node, BinaryOpNode):
            walk(node.left)
            walk(node.right)
        elif isinstance(node, UnaryOpNode):
            walk(node.operand)

    walk(ast)
    return count


def function_skeleton(ast: AstNode) -> str:
    """Return a preorder function-name skeleton for ``ast``."""
    names: list[str] = []

    def walk(node: AstNode) -> None:
        if isinstance(node, FunctionCallNode):
            names.append(node.name)
            for arg in node.args:
                walk(arg)
        elif isinstance(node, BinaryOpNode):
            walk(node.left)
            walk(node.right)
        elif isinstance(node, UnaryOpNode):
            walk(node.operand)

    walk(ast)
    return ">".join(names) if names else "leaf"


def structural_signature(
    root: NodeKey,
    formula: str,
    graph: DependencyGraph,
) -> StructuralSignature:
    """Build a cheap structural signature for one collapsed root."""
    del graph
    row, _column = _root_row_column(root)
    stripped = formula.strip()
    if stripped.startswith("="):
        stripped = stripped[1:]
    try:
        ast = parse(stripped)
    except FormulaParseError:
        return StructuralSignature(
            kind=classify_formula_kind(formula),
            row_band=row,
            dependency_count=0,
            function_skeleton="unparsed",
        )
    return StructuralSignature(
        kind=classify_formula_kind(formula),
        row_band=row,
        dependency_count=_dependency_count(ast),
        function_skeleton=function_skeleton(ast),
    )
