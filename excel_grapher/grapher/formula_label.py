"""Shared formula display helpers for graph export and lightweight visualization."""


def validate_max_formula_length(max_formula_length: int | None) -> None:
    if max_formula_length is None:
        return
    if max_formula_length <= 0:
        raise ValueError("max_formula_length must be None or a positive integer")


def truncate_formula_display(formula: str, max_formula_length: int | None) -> str:
    if max_formula_length is None or len(formula) <= max_formula_length:
        return formula
    return formula[:max_formula_length] + "..."
