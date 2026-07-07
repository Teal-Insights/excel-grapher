"""Excel error exceptions and raise helpers for the exported Python runtime."""

from __future__ import annotations

from typing import NoReturn

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.types import XlErrorException

__all__ = ["XlErrorException", "raise_if_sentinel", "xl_raise"]


def xl_raise(code: XlError) -> NoReturn:
    """Raise an Excel error code from an expression position."""
    raise XlErrorException(code)


def raise_if_sentinel(value: CellValue) -> CellValue:
    """Return *value* or raise ``XlErrorException`` when it is an Excel error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value
