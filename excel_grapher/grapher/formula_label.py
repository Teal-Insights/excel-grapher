"""Shared formula display helpers for graph export and lightweight visualization."""

from __future__ import annotations

from typing import Protocol


class _HasFormulas(Protocol):
    formula: str | None
    normalized_formula: str | None


def display_formula(node: _HasFormulas) -> str | None:
    """Return the best formula text to show for `node`.

    Prefers the raw workbook string, which extraction stores only when
    `store_raw_formula=True`, and otherwise falls back to `normalized_formula`
    so visualizations stay informative on graphs built without it.
    """
    return node.formula if node.formula is not None else node.normalized_formula


def validate_max_formula_length(max_formula_length: int | None) -> None:
    if max_formula_length is None:
        return
    if max_formula_length <= 0:
        raise ValueError("max_formula_length must be None or a positive integer")


def truncate_formula_display(formula: str, max_formula_length: int | None) -> str:
    if max_formula_length is None or len(formula) <= max_formula_length:
        return formula
    return formula[:max_formula_length] + "..."
