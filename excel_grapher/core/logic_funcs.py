"""Shared logical worksheet function implementations."""

from __future__ import annotations

from .coercions import get_error, to_bool
from .types import CellValue, XlError

__all__ = ["logical_and", "logical_not", "logical_or"]


def logical_and(*args: CellValue) -> bool | XlError:
    """Return logical AND across arguments."""
    err = get_error(*args)
    if err is not None:
        return err
    for a in args:
        b = to_bool(a)
        if isinstance(b, XlError):
            return b
        if not b:
            return False
    return True


def logical_or(*args: CellValue) -> bool | XlError:
    """Return logical OR across arguments."""
    err = get_error(*args)
    if err is not None:
        return err
    for a in args:
        b = to_bool(a)
        if isinstance(b, XlError):
            return b
        if b:
            return True
    return False


def logical_not(arg: CellValue) -> bool | XlError:
    """Return logical NOT of an argument."""
    b = to_bool(arg)
    if isinstance(b, XlError):
        return b
    return not b
