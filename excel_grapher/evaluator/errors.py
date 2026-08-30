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
    """Raised when a formula cell lacks `formula_ast` and `normalized_formula`."""

    def __init__(self, cell_key: str) -> None:
        super().__init__(
            f"Cell {cell_key!r} has a formula but formula_ast and normalized_formula "
            "are missing; rebuild the dependency graph or set formula_ast."
        )
        self.cell_key = cell_key
