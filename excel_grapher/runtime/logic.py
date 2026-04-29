from __future__ import annotations

from excel_grapher.core import CellValue, XlError, get_error, to_bool, to_number

__all__ = ["xl_and", "xl_choose", "xl_ifna", "xl_not", "xl_or"]


def xl_and(*args: CellValue) -> bool | XlError:
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


def xl_or(*args: CellValue) -> bool | XlError:
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


def xl_not(arg: CellValue) -> bool | XlError:
    b = to_bool(arg)
    if isinstance(b, XlError):
        return b
    return not b


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
