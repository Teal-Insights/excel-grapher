"""Raise-only boundary wrappers for logic worksheet functions."""

from __future__ import annotations

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.logic_funcs import logical_and, logical_if, logical_not, logical_or
from excel_grapher.core.types import XlErrorException

from .errors import raise_if_sentinel_bool

__all__ = ["xl_and", "xl_if", "xl_not", "xl_or"]


def xl_and(*args: CellValue) -> bool:
    """Return logical AND, raising on Excel errors."""
    return raise_if_sentinel_bool(logical_and(*args))


def xl_or(*args: CellValue) -> bool:
    """Return logical OR, raising on Excel errors."""
    return raise_if_sentinel_bool(logical_or(*args))


def xl_not(arg: CellValue) -> bool:
    """Return logical NOT, raising on Excel errors."""
    return raise_if_sentinel_bool(logical_not(arg))


def xl_if(
    cond: object,
    then_value: object,
    else_value: object = False,
) -> object:
    """Excel `IF` via `logical_if` (scalar or element-wise).

    A whole-result `#VALUE!` (shape mismatch) raises. Per-cell errors stay in
    the returned grid for aggregates to consume.
    """
    result = logical_if(cond, then_value, else_value)
    if isinstance(result, XlError):
        raise XlErrorException(result)
    return result
