from __future__ import annotations

from .core import CellValue, XlError, to_bool, to_number


def xl_and(*args: CellValue) -> bool | XlError:
    for a in args:
        b = to_bool(a)
        if isinstance(b, XlError):
            return b
        if not b:
            return False
    return True


def xl_or(*args: CellValue) -> bool | XlError:
    any_true = False
    for a in args:
        b = to_bool(a)
        if isinstance(b, XlError):
            return b
        any_true = any_true or b
    return any_true


def xl_choose(index_num: CellValue, *values: CellValue) -> CellValue:
    idx = to_number(index_num)
    if isinstance(idx, XlError):
        return idx
    i = int(idx)
    if i < 1 or i > len(values):
        return XlError.VALUE
    return values[i - 1]


def xl_ifna(value: CellValue, value_if_na: CellValue) -> CellValue:
    if value == XlError.NA:
        return value_if_na
    return value

