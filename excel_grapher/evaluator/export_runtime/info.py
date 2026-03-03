from __future__ import annotations

from .core import CellValue, XlError, to_number


def xl_isnumber(value: CellValue) -> bool:
    n = to_number(value)
    return isinstance(n, float) and not isinstance(value, bool)


def xl_istext(value: CellValue) -> bool:
    return isinstance(value, str)


def xl_isblank(value: CellValue) -> bool:
    return value is None


def xl_na() -> XlError:
    return XlError.NA

