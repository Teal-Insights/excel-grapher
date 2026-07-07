"""Excel error exceptions and raise helpers for the exported Python runtime."""

from __future__ import annotations

from typing import NoReturn

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.types import XlErrorException

__all__ = [
    "XlErrorException",
    "raise_if_sentinel_bool",
    "raise_if_sentinel_float",
    "raise_if_sentinel_int",
    "raise_if_sentinel_str",
    "xl_raise",
]


def xl_raise(code: XlError) -> NoReturn:
    """Raise an Excel error code from an expression position."""
    raise XlErrorException(code)


def raise_if_sentinel_float(value: float | XlError) -> float:
    """Return a float result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value


def raise_if_sentinel_int(value: int | XlError) -> int:
    """Return an integer result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value


def raise_if_sentinel_str(value: str | XlError) -> str:
    """Return a string result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value


def raise_if_sentinel_bool(value: bool | XlError) -> bool:
    """Return a boolean result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value


def raise_if_sentinel(value: CellValue) -> CellValue:
    """Return *value* or raise ``XlErrorException`` when it is an Excel error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value
