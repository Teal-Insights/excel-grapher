from __future__ import annotations

import numpy as np

from .core import CellValue, XlError

__all__ = ["xl_isblank", "xl_isnumber", "xl_istext", "xl_na"]


def xl_isnumber(value: CellValue) -> bool:
    return not isinstance(value, bool) and isinstance(value, (int, float, np.integer, np.floating))


def xl_istext(value: CellValue) -> bool:
    return isinstance(value, str)


def xl_isblank(value: CellValue) -> bool:
    return value is None


def xl_na() -> XlError:
    return XlError.NA
