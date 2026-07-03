"""Lookup and reference functions over lazy ranges for exported code."""

from __future__ import annotations

from excel_grapher.core import XlError, excel_casefold, to_number
from excel_grapher.core.types import XlErrorException

from .values import CellValue, Grid, Scalar, as_scalar

__all__ = [
    "xl_hlookup",
    "xl_index",
    "xl_lookup",
    "xl_match",
    "xl_vlookup",
    "xl_xlookup",
]


def _number_arg(value: CellValue) -> float:
    """Coerce a scalar function argument, raising on Excel coercion errors."""
    number = to_number(as_scalar(value))
    if isinstance(number, XlError):
        raise XlErrorException(number)
    return number


def _result_or_raise(value: Scalar) -> Scalar:
    """Return a lookup result cell, raising when it holds an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value


def _values_match(a: CellValue, b: CellValue) -> bool:
    a = as_scalar(a)
    b = as_scalar(b)
    if isinstance(a, str) and isinstance(b, str):
        return excel_casefold(a) == excel_casefold(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return an == bn
    return a == b


def _compare_values(a: CellValue, b: CellValue) -> int:
    a = as_scalar(a)
    b = as_scalar(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return -1 if an < bn else 1 if an > bn else 0
    if isinstance(a, str) and isinstance(b, str):
        af = excel_casefold(a)
        bf = excel_casefold(b)
        return -1 if af < bf else 1 if af > bf else 0
    return 0


def _vector_of(grid: Grid) -> Grid | None:
    """Return the grid when it is a single row or column vector."""
    if grid.nrows == 1 or grid.ncols == 1:
        return grid
    return None


def xl_lookup(
    lookup_value: CellValue,
    lookup_vector_or_array: CellValue,
    result_vector: CellValue = None,
) -> CellValue:
    grid = Grid.wrap(lookup_vector_or_array)
    if grid is None:
        raise XlErrorException(XlError.VALUE)
    result_grid = Grid.wrap(result_vector) if result_vector is not None else None
    if result_vector is not None and result_grid is None:
        raise XlErrorException(XlError.VALUE)

    if result_grid is None:
        vector = _vector_of(grid)
        if vector is not None:
            lookup_flat = vector
            result_flat = vector
        elif grid.nrows >= grid.ncols:
            lookup_flat = Grid.wrap(grid.col_slice(0))
            result_flat = Grid.wrap(grid.col_slice(grid.ncols - 1))
            assert lookup_flat is not None and result_flat is not None
        else:
            lookup_flat = Grid.wrap(grid.row_slice(0))
            result_flat = Grid.wrap(grid.row_slice(grid.nrows - 1))
            assert lookup_flat is not None and result_flat is not None
    else:
        if _vector_of(grid) is None or _vector_of(result_grid) is None:
            raise XlErrorException(XlError.NA)
        if grid.size != result_grid.size:
            raise XlErrorException(XlError.NA)
        lookup_flat = grid
        result_flat = result_grid

    last_match_idx = None
    for i in range(lookup_flat.size):
        if _compare_values(lookup_flat.at_flat(i), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        raise XlErrorException(XlError.NA)
    return _result_or_raise(result_flat.at_flat(last_match_idx))


def xl_index(array: CellValue, row_num: CellValue, col_num: CellValue = None) -> CellValue:
    grid = Grid.wrap(array)
    if grid is None:
        raise XlErrorException(XlError.VALUE)
    nrows, ncols = grid.nrows, grid.ncols
    row_omitted = row_num is None
    col_omitted = col_num is None

    if row_omitted and col_omitted:
        if nrows == 1 and ncols == 1:
            return _result_or_raise(grid.at(0, 0))
        if nrows == 1:
            return _result_or_raise(grid.at(0, ncols - 1))
        if ncols == 1:
            return _result_or_raise(grid.at(nrows - 1, 0))
        raise XlErrorException(XlError.VALUE)

    if row_omitted:
        cn = _number_arg(col_num)
        col = int(cn)
        if col < 1 or col > ncols:
            raise XlErrorException(XlError.REF)
        if nrows == 1:
            return _result_or_raise(grid.at(0, col - 1))
        return grid.col_slice(col - 1)

    rn = _number_arg(row_num)
    row = int(rn)

    if col_omitted:
        if nrows == 1:
            if row < 1 or row > ncols:
                raise XlErrorException(XlError.REF)
            return _result_or_raise(grid.at(0, row - 1))
        if ncols == 1:
            if row < 1 or row > nrows:
                raise XlErrorException(XlError.REF)
            return _result_or_raise(grid.at(row - 1, 0))
        if row < 1 or row > nrows:
            raise XlErrorException(XlError.REF)
        return grid.row_slice(row - 1)

    cn = _number_arg(col_num)
    col = int(cn)
    if nrows == 1:
        if row < 1 or row > ncols:
            raise XlErrorException(XlError.REF)
        return _result_or_raise(grid.at(0, row - 1))
    if ncols == 1:
        if row < 1 or row > nrows:
            raise XlErrorException(XlError.REF)
        return _result_or_raise(grid.at(row - 1, 0))
    if row < 1 or row > nrows:
        raise XlErrorException(XlError.REF)
    if col < 1 or col > ncols:
        raise XlErrorException(XlError.REF)
    return _result_or_raise(grid.at(row - 1, col - 1))


def xl_match(lookup_value: CellValue, lookup_array: CellValue, match_type: CellValue = 1) -> int:
    mt = _number_arg(match_type)
    match_type_int = int(mt)
    if isinstance(lookup_array, XlError):
        raise XlErrorException(lookup_array)
    grid = Grid.wrap(lookup_array)
    if grid is None:
        grid_wrapped = Grid.wrap([[lookup_array]])
        assert grid_wrapped is not None
        grid = grid_wrapped
    if match_type_int == 0:
        for i in range(grid.size):
            if _values_match(lookup_value, grid.at_flat(i)):
                return i + 1
        raise XlErrorException(XlError.NA)
    if match_type_int == 1:
        last_match = None
        for i in range(grid.size):
            if _compare_values(grid.at_flat(i), lookup_value) <= 0:
                last_match = i + 1
            else:
                break
        if last_match is None:
            raise XlErrorException(XlError.NA)
        return last_match
    if match_type_int == -1:
        last_match = None
        for i in range(grid.size):
            if _compare_values(grid.at_flat(i), lookup_value) >= 0:
                last_match = i + 1
            else:
                break
        if last_match is None:
            raise XlErrorException(XlError.NA)
        return last_match
    raise XlErrorException(XlError.VALUE)


def xl_vlookup(
    lookup_value: CellValue,
    table_array: CellValue,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    cn = _number_arg(col_index_num)
    col_index = int(cn)
    if col_index < 1:
        raise XlErrorException(XlError.VALUE)
    grid = Grid.wrap(table_array)
    if grid is None:
        raise XlErrorException(XlError.VALUE)
    if col_index > grid.ncols:
        raise XlErrorException(XlError.REF)
    exact_match = not bool(range_lookup)
    if exact_match:
        for i in range(grid.nrows):
            if _values_match(lookup_value, grid.at(i, 0)):
                return _result_or_raise(grid.at(i, col_index - 1))
        raise XlErrorException(XlError.NA)
    last_match_idx = None
    for i in range(grid.nrows):
        if _compare_values(grid.at(i, 0), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        raise XlErrorException(XlError.NA)
    return _result_or_raise(grid.at(last_match_idx, col_index - 1))


def xl_hlookup(
    lookup_value: CellValue,
    table_array: CellValue,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    rn = _number_arg(row_index_num)
    row_index = int(rn)
    if row_index < 1:
        raise XlErrorException(XlError.VALUE)
    grid = Grid.wrap(table_array)
    if grid is None:
        raise XlErrorException(XlError.VALUE)
    if row_index > grid.nrows:
        raise XlErrorException(XlError.REF)
    exact_match = not bool(range_lookup)
    if exact_match:
        for i in range(grid.ncols):
            if _values_match(lookup_value, grid.at(0, i)):
                return _result_or_raise(grid.at(row_index - 1, i))
        raise XlErrorException(XlError.NA)
    last_match_idx = None
    for i in range(grid.ncols):
        if _compare_values(grid.at(0, i), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        raise XlErrorException(XlError.NA)
    return _result_or_raise(grid.at(row_index - 1, last_match_idx))


def xl_xlookup(
    lookup_value: CellValue,
    lookup_array: CellValue,
    return_array: CellValue,
    if_not_found: CellValue = None,
    match_mode: CellValue = 0,
    search_mode: CellValue = 1,
) -> CellValue:
    """Excel XLOOKUP.

    This implementation supports:
    - exact match (match_mode=0)
    - search first-to-last (search_mode=1) and last-to-first (search_mode=-1)
    """
    mm = _number_arg(match_mode)
    sm = _number_arg(search_mode)

    mm_i = int(mm)
    sm_i = int(sm)

    if mm_i != 0:
        raise XlErrorException(XlError.VALUE)
    if sm_i not in (1, -1):
        raise XlErrorException(XlError.VALUE)

    keys = Grid.wrap(lookup_array)
    vals = Grid.wrap(return_array)
    if keys is None or vals is None:
        raise XlErrorException(XlError.VALUE)
    if keys.size != vals.size:
        raise XlErrorException(XlError.VALUE)

    idxs = range(keys.size) if sm_i == 1 else range(keys.size - 1, -1, -1)
    for i in idxs:
        if _values_match(lookup_value, keys.at_flat(i)):
            return _result_or_raise(vals.at_flat(i))

    if if_not_found is None:
        raise XlErrorException(XlError.NA)
    return if_not_found
