from __future__ import annotations


class FormulaExpanderError(Exception):
    """Base exception for excel-formula-expander failures."""


class ParseError(FormulaExpanderError):
    """Raised when a formula cannot be parsed into an AST."""

    def __init__(self, formula: str, message: str) -> None:
        super().__init__(f"Parse error: {message}. Formula: {formula!r}")
        self.formula = formula
        self.message = message


class MissingNormalizedFormulaError(FormulaExpanderError):
    """Raised when a formula cell lacks `normalized_formula` (graph invariant)."""

    def __init__(self, cell_key: str) -> None:
        super().__init__(
            f"Cell {cell_key!r} has a formula but normalized_formula is missing; "
            "rebuild the dependency graph or set normalized_formula."
        )
        self.cell_key = cell_key


class FormulaGroupKeyError(FormulaExpanderError):
    """Raised when evaluating a multi-cell group key (vector eval is out of scope)."""

    def __init__(self, group_key: str) -> None:
        super().__init__(
            f"Cannot evaluate multi-cell group key {group_key!r}; "
            "evaluate a member cell address instead (Option B)."
        )
        self.group_key = group_key


class MissingGroupTemplateError(FormulaExpanderError):
    """Raised when a formula-group node lacks skeleton/bindings for a member."""

    def __init__(self, group_key: str, member_key: str) -> None:
        super().__init__(
            f"Formula-group node {group_key!r} has no usable template for member "
            f"{member_key!r}; set skeleton and member_bindings."
        )
        self.group_key = group_key
        self.member_key = member_key
