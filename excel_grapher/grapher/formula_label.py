"""Shared formula display helpers for graph export and lightweight visualization."""

from __future__ import annotations

from typing import Protocol

from excel_grapher.core.formula_ast import unparse_normalized_formula


class _HasFormulas(Protocol):
    formula: str | None
    normalized_formula: str | None


def display_formula(node: _HasFormulas) -> str | None:
    """Return the best formula text to show for `node`.

    Prefers the raw workbook string, which extraction stores only when
    `store_raw_formula=True`. Otherwise renders `formula_ast` as absolute A1,
    then falls back to stored `normalized_formula` for unparseable cells.
    """
    if node.formula is not None:
        return node.formula
    ast = getattr(node, "formula_ast", None)
    if ast is not None:
        anchor = getattr(node, "address", None)
        if anchor is None:
            anchor = getattr(node, "key", None)
        return unparse_normalized_formula(ast, anchor=anchor)
    return node.normalized_formula


def validate_max_formula_length(max_formula_length: int | None) -> None:
    if max_formula_length is None:
        return
    if max_formula_length <= 0:
        raise ValueError("max_formula_length must be None or a positive integer")


def truncate_formula_display(formula: str, max_formula_length: int | None) -> str:
    if max_formula_length is None or len(formula) <= max_formula_length:
        return formula
    return formula[:max_formula_length] + "..."
