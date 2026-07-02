"""Excel error exceptions and raise helpers for the exported Python runtime."""

from __future__ import annotations

from typing import NoReturn

from excel_grapher.core import XlError
from excel_grapher.core.types import XlErrorException

__all__ = ["XlErrorException", "xl_raise"]


def xl_raise(code: XlError) -> NoReturn:
    """Raise an Excel error code from an expression position."""
    raise XlErrorException(code)
