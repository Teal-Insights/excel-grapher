"""Error-consuming functions for exported code (raise-based error channel).

Exported expressions raise `XlErrorException`, so error-consuming functions
receive lazily-evaluated thunks instead of pre-evaluated values. Each helper
also handles `XlError` sentinel returns from embedded shared functions that
still use the sentinel channel internally.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import cast

from excel_grapher.core import XlError
from excel_grapher.core.types import XlErrorException

from .ranges import Range
from .values import CellValue

__all__ = [
    "xl_iferror",
    "xl_ifna",
    "xl_isblank",
    "xl_iserror",
    "xl_isna",
]


def _resolve_scalar(value_fn: Callable[[], CellValue]) -> CellValue:
    """Evaluate a thunk, resolving 1x1 range views to their single cell value."""
    value = value_fn()
    if isinstance(value, Range) and value.shape == (1, 1):
        return cast("CellValue", value.cell(1, 1))
    return value


def xl_iferror(
    value_fn: Callable[[], CellValue], fallback_fn: Callable[[], CellValue]
) -> CellValue:
    """Excel IFERROR over lazily-evaluated value and fallback thunks."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException:
        return fallback_fn()
    if isinstance(value, XlError):
        return fallback_fn()
    return value


def xl_ifna(value_fn: Callable[[], CellValue], fallback_fn: Callable[[], CellValue]) -> CellValue:
    """Excel IFNA: catch `#N/A` only; other Excel errors propagate."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException as exc:
        if exc.code == XlError.NA:
            return fallback_fn()
        raise
    if value == XlError.NA:
        return fallback_fn()
    return value


def xl_iserror(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISERROR: True when evaluating the argument produces any Excel error."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException:
        return True
    return isinstance(value, XlError)


def xl_isna(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISNA: True when evaluating the argument produces `#N/A`."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException as exc:
        return exc.code == XlError.NA
    return value == XlError.NA


def xl_isblank(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISBLANK: IS functions do not propagate errors."""
    try:
        value = value_fn()
    except XlErrorException:
        return False
    if isinstance(value, Range):
        if value.shape != (1, 1):
            return False
        value = value.value_at(1, 1)
    return value is None
