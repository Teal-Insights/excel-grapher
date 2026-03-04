from __future__ import annotations

from ..export_runtime.logic import xl_ifna as _rt_xl_ifna
from ..helpers import get_error, to_bool, to_number
from ..types import CellValue, XlError
from . import register


@register("AND")
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


@register("OR")
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


@register("CHOOSE")
def xl_choose(index_num: CellValue, *values: CellValue) -> CellValue:
    """Excel CHOOSE function - returns a value from a list based on index.

    CHOOSE(index_num, value1, [value2], ...)

    Args:
        index_num: The index (1-based) of the value to return.
        values: The values to choose from.

    Returns:
        The value at the specified index, or an error if invalid.
    """
    # Handle error in index
    if isinstance(index_num, XlError):
        return index_num

    n = to_number(index_num)
    if isinstance(n, XlError):
        return n
    idx = int(n)

    # Check bounds (1-based index)
    if idx < 1 or idx > len(values):
        return XlError.VALUE

    return values[idx - 1]


@register("_XLFN.IFNA")
@register("IFNA")
def xl_ifna(value: CellValue, value_if_na: CellValue) -> CellValue:
    """Return value_if_na if value is #N/A; otherwise return value."""
    return _rt_xl_ifna(value, value_if_na)
