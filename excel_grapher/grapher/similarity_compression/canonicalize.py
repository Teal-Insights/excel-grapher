"""Canonical text blobs for embedding collapsed formulas."""

from __future__ import annotations

import re
from collections.abc import Iterable

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
from .signatures import function_skeleton, structural_signature

__all__ = ["canonicalize_for_embedding"]


def _root_row_column(root: NodeKey) -> tuple[int, str]:
    _, rest = root.split("!", 1)
    column = "".join(char for char in rest if char.isalpha())
    row = int("".join(char for char in rest if char.isdigit()))
    return row, column.upper()


def _normalize_address_roles(text: str, *, root_column: str) -> str:
    col = re.escape(root_column.upper())
    out = text
    out = re.sub(rf"(Engine!){col}(\d+)", r"\1{COL}\2", out, flags=re.IGNORECASE)
    out = re.sub(r"Engine!B(\d+)", r"Engine!{BASE}\1", out, flags=re.IGNORECASE)
    out = re.sub(rf"(Inputs!){col}(\d+)", r"\1{COL}\2", out, flags=re.IGNORECASE)
    return out


def _role_normalize_formula(formula: str, *, root_column: str) -> str:
    return _normalize_address_roles(formula, root_column=root_column)


def _external_dependencies(ast: AstNode) -> tuple[str, ...]:
    deps: set[str] = set()

    def walk(node: AstNode) -> None:
        if isinstance(node, CellRefNode):
            deps.add(node.address)
        elif isinstance(node, RangeNode):
            deps.add(node.start)
            deps.add(node.end)
        elif isinstance(node, FunctionCallNode):
            for arg in node.args:
                walk(arg)
        elif isinstance(node, BinaryOpNode):
            walk(node.left)
            walk(node.right)
        elif isinstance(node, UnaryOpNode):
            walk(node.operand)

    walk(ast)
    return tuple(sorted(deps))


def _format_blob(
    *,
    kind: str,
    root: NodeKey,
    row: int,
    column: str,
    formula_normalized: str,
    external_dependencies: Iterable[str],
    skeleton: str,
) -> str:
    lines = [
        f"kind: {kind}",
        f"root: {_normalize_address_roles(root, root_column=column)}",
        f"row_labels: row={row}",
        "column_labels: column_role={COL}",
        f"formula_normalized: {formula_normalized}",
        f"function_skeleton: {skeleton}",
        "external_dependencies: "
        + ", ".join(
            _normalize_address_roles(dep, root_column=column) for dep in external_dependencies
        ),
    ]
    return "\n".join(lines)


def canonicalize_for_embedding(
    root: NodeKey,
    formula: str,
    graph: DependencyGraph,
) -> str:
    """Build a structured embedding input for one collapsed root formula.

    Args:
        root: Retained computation address.
        formula: Collapsed normalized formula text.
        graph: Canonical graph (reserved for future label injection).

    Returns:
        Multi-line text blob suitable for generalist embedding models.
    """
    row, column = _root_row_column(root)
    signature = structural_signature(root, formula, graph)
    stripped = formula.strip()
    if stripped.startswith("="):
        stripped = stripped[1:]
    role_formula = _role_normalize_formula(stripped, root_column=column)
    try:
        ast = parse(stripped)
        externals = _external_dependencies(ast)
        skeleton = function_skeleton(ast)
    except FormulaParseError:
        externals = ()
        skeleton = signature.function_skeleton
    return _format_blob(
        kind=signature.kind,
        root=root,
        row=row,
        column=column,
        formula_normalized=role_formula,
        external_dependencies=externals,
        skeleton=skeleton,
    )
