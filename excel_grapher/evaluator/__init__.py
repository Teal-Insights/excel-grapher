"""
Formula Expander: evaluate Excel formulas from an `excel_grapher.DependencyGraph`.

The public API is intentionally small and stable; internal modules may change.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from .errors import ParseError
from .types import CellValue, ExcelRange, XlError

if TYPE_CHECKING:  # pragma: no cover
    from .evaluator import FormulaEvaluator

__all__ = [
    "FormulaEvaluator",
    "CellValue",
    "ExcelRange",
    "XlError",
    "ParseError",
]


def __getattr__(name: str):
    # Lazy import to keep the package importable while modules are developed.
    if name == "FormulaEvaluator":
        from .evaluator import FormulaEvaluator

        return FormulaEvaluator
    raise AttributeError(name)
