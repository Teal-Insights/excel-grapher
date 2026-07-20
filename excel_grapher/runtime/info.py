from __future__ import annotations

import numbers

from excel_grapher.core import CellValue, XlError

__all__ = ["xl_isblank", "xl_iserror", "xl_isna", "xl_isnumber", "xl_istext", "xl_na"]


def xl_isnumber(value: CellValue) -> bool:
    return not isinstance(value, bool) and isinstance(value, numbers.Real)


def xl_istext(value: CellValue) -> bool:
    # XlError subclasses str; Excel ISTEXT returns FALSE for error values.
    return isinstance(value, str) and not isinstance(value, XlError)


def xl_isblank(value: CellValue) -> bool:
    return value is None


def xl_na() -> XlError:
    return XlError.NA


def xl_iserror(value: CellValue) -> bool:
    return isinstance(value, XlError)


def xl_isna(value: CellValue) -> bool:
    return value == XlError.NA
