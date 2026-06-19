"""Excel function semantics shared by expr_eval, runtime, and export."""

from __future__ import annotations

from .coercions import to_number
from .types import CellValue, XlError

__all__ = ["xl_abs"]


def xl_abs(*args: CellValue) -> float | XlError:
    """Return the absolute value of a number (Excel ``ABS``)."""
    if len(args) != 1:
        return XlError.VALUE
    n = to_number(args[0])
    if isinstance(n, XlError):
        return n
    return float(abs(n))
