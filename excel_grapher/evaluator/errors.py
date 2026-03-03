from __future__ import annotations


class FormulaExpanderError(Exception):
    """Base exception for excel-formula-expander failures."""


class ParseError(FormulaExpanderError):
    """Raised when a formula cannot be parsed into an AST."""

    def __init__(self, formula: str, message: str) -> None:
        super().__init__(f"Parse error: {message}. Formula: {formula!r}")
        self.formula = formula
        self.message = message
