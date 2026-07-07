"""Raise-only boundary wrappers for logic worksheet functions."""

from __future__ import annotations

from excel_grapher.core import CellValue
from excel_grapher.core.logic_funcs import logical_and, logical_not, logical_or

from .errors import raise_if_sentinel_bool

__all__ = ["xl_and", "xl_not", "xl_or"]


def xl_and(*args: CellValue) -> bool:
    """Return logical AND, raising on Excel errors."""
    return raise_if_sentinel_bool(logical_and(*args))


def xl_or(*args: CellValue) -> bool:
    """Return logical OR, raising on Excel errors."""
    return raise_if_sentinel_bool(logical_or(*args))


def xl_not(arg: CellValue) -> bool:
    """Return logical NOT, raising on Excel errors."""
    return raise_if_sentinel_bool(logical_not(arg))
