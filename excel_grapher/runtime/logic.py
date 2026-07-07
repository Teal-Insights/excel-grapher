from __future__ import annotations

from excel_grapher.core import CellValue, XlError, to_number
from excel_grapher.core.logic_funcs import logical_and, logical_not, logical_or

__all__ = ["xl_and", "xl_choose", "xl_ifna", "xl_not", "xl_or"]


def xl_and(*args: CellValue) -> bool | XlError:
    return logical_and(*args)


def xl_or(*args: CellValue) -> bool | XlError:
    return logical_or(*args)


def xl_not(arg: CellValue) -> bool | XlError:
    return logical_not(arg)


def xl_choose(index_num: CellValue, *values: CellValue) -> CellValue:
    if isinstance(index_num, XlError):
        return index_num
    n = to_number(index_num)
    if isinstance(n, XlError):
        return n
    idx = int(n)
    if idx < 1 or idx > len(values):
        return XlError.VALUE
    return values[idx - 1]


def xl_ifna(value: CellValue, value_if_na: CellValue) -> CellValue:
    if value == XlError.NA:
        return value_if_na
    return value
