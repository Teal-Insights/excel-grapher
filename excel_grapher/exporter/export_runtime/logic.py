"""Raise-only boundary wrappers for logic runtime helpers."""

from __future__ import annotations

from excel_grapher.core import CellValue

from .errors import raise_if_sentinel_bool

__all__ = ["xl_and", "xl_not", "xl_or"]


def xl_and(*args: CellValue) -> bool:
    """Return logical AND, raising on Excel errors."""
    return raise_if_sentinel_bool(_sentinel_xl_and(*args))


def xl_or(*args: CellValue) -> bool:
    """Return logical OR, raising on Excel errors."""
    return raise_if_sentinel_bool(_sentinel_xl_or(*args))


def xl_not(arg: CellValue) -> bool:
    """Return logical NOT, raising on Excel errors."""
    return raise_if_sentinel_bool(_sentinel_xl_not(arg))
