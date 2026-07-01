"""Lookup and reference functions over lazy ranges for exported code."""

from __future__ import annotations

from excel_grapher.core import XlError, excel_casefold, to_number

from .values import CellValue, Grid, as_scalar

__all__ = [
    "xl_hlookup",
    "xl_index",
    "xl_lookup",
    "xl_match",
    "xl_vlookup",
    "xl_xlookup",
]


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
        return XlError.VALUE
    result_grid = Grid.wrap(result_vector) if result_vector is not None else None
    if result_vector is not None and result_grid is None:
        return XlError.VALUE

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
            return XlError.NA
        if grid.size != result_grid.size:
            return XlError.NA
        lookup_flat = grid
        result_flat = result_grid

    last_match_idx = None
    for i in range(lookup_flat.size):
        if _compare_values(lookup_flat.at_flat(i), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return result_flat.at_flat(last_match_idx)


def xl_index(array: CellValue, row_num: CellValue, col_num: CellValue = None) -> CellValue:
    grid = Grid.wrap(array)
    if grid is None:
        return XlError.VALUE
    nrows, ncols = grid.nrows, grid.ncols
    row_omitted = row_num is None
    col_omitted = col_num is None

    if row_omitted and col_omitted:
        if nrows == 1 and ncols == 1:
            return grid.at(0, 0)
        if nrows == 1:
            return grid.at(0, ncols - 1)
        if ncols == 1:
            return grid.at(nrows - 1, 0)
        return XlError.VALUE

    if row_omitted:
        cn = to_number(as_scalar(col_num))
        if isinstance(cn, XlError):
            return cn
        col = int(cn)
        if col < 1 or col > ncols:
            return XlError.REF
        if nrows == 1:
            return grid.at(0, col - 1)
        return grid.col_slice(col - 1)

    rn = to_number(as_scalar(row_num))
    if isinstance(rn, XlError):
        return rn
    row = int(rn)

    if col_omitted:
        if nrows == 1:
            if row < 1 or row > ncols:
                return XlError.REF
            return grid.at(0, row - 1)
        if ncols == 1:
            if row < 1 or row > nrows:
                return XlError.REF
            return grid.at(row - 1, 0)
        if row < 1 or row > nrows:
            return XlError.REF
        return grid.row_slice(row - 1)

    cn = to_number(as_scalar(col_num))
    if isinstance(cn, XlError):
        return cn
    col = int(cn)
    if nrows == 1:
        if row < 1 or row > ncols:
            return XlError.REF
        return grid.at(0, row - 1)
    if ncols == 1:
        if row < 1 or row > nrows:
            return XlError.REF
        return grid.at(row - 1, 0)
    if row < 1 or row > nrows:
        return XlError.REF
    if col < 1 or col > ncols:
        return XlError.REF
    return grid.at(row - 1, col - 1)


def xl_match(
    lookup_value: CellValue, lookup_array: CellValue, match_type: CellValue = 1
) -> int | XlError:
    mt = to_number(as_scalar(match_type))
    if isinstance(mt, XlError):
        return mt
    match_type_int = int(mt)
    if isinstance(lookup_array, XlError):
        return lookup_array
    grid = Grid.wrap(lookup_array)
    if grid is None:
        grid_wrapped = Grid.wrap([[lookup_array]])
        assert grid_wrapped is not None
        grid = grid_wrapped
    if match_type_int == 0:
        for i in range(grid.size):
            if _values_match(lookup_value, grid.at_flat(i)):
                return i + 1
        return XlError.NA
    if match_type_int == 1:
        last_match = None
        for i in range(grid.size):
            if _compare_values(grid.at_flat(i), lookup_value) <= 0:
                last_match = i + 1
            else:
                break
        return XlError.NA if last_match is None else last_match
    if match_type_int == -1:
        last_match = None
        for i in range(grid.size):
            if _compare_values(grid.at_flat(i), lookup_value) >= 0:
                last_match = i + 1
            else:
                break
        return XlError.NA if last_match is None else last_match
    return XlError.VALUE


def xl_vlookup(
    lookup_value: CellValue,
    table_array: CellValue,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    cn = to_number(as_scalar(col_index_num))
    if isinstance(cn, XlError):
        return cn
    col_index = int(cn)
    if col_index < 1:
        return XlError.VALUE
    grid = Grid.wrap(table_array)
    if grid is None:
        return XlError.VALUE
    if col_index > grid.ncols:
        return XlError.REF
    exact_match = not bool(range_lookup)
    if exact_match:
        for i in range(grid.nrows):
            if _values_match(lookup_value, grid.at(i, 0)):
                return grid.at(i, col_index - 1)
        return XlError.NA
    last_match_idx = None
    for i in range(grid.nrows):
        if _compare_values(grid.at(i, 0), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return grid.at(last_match_idx, col_index - 1)


def xl_hlookup(
    lookup_value: CellValue,
    table_array: CellValue,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    rn = to_number(as_scalar(row_index_num))
    if isinstance(rn, XlError):
        return rn
    row_index = int(rn)
    if row_index < 1:
        return XlError.VALUE
    grid = Grid.wrap(table_array)
    if grid is None:
        return XlError.VALUE
    if row_index > grid.nrows:
        return XlError.REF
    exact_match = not bool(range_lookup)
    if exact_match:
        for i in range(grid.ncols):
            if _values_match(lookup_value, grid.at(0, i)):
                return grid.at(row_index - 1, i)
        return XlError.NA
    last_match_idx = None
    for i in range(grid.ncols):
        if _compare_values(grid.at(0, i), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return grid.at(row_index - 1, last_match_idx)


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
    mm = to_number(as_scalar(match_mode))
    if isinstance(mm, XlError):
        return mm
    sm = to_number(as_scalar(search_mode))
    if isinstance(sm, XlError):
        return sm

    mm_i = int(mm)
    sm_i = int(sm)

    if mm_i != 0:
        return XlError.VALUE
    if sm_i not in (1, -1):
        return XlError.VALUE

    keys = Grid.wrap(lookup_array)
    vals = Grid.wrap(return_array)
    if keys is None or vals is None:
        return XlError.VALUE
    if keys.size != vals.size:
        return XlError.VALUE

    idxs = range(keys.size) if sm_i == 1 else range(keys.size - 1, -1, -1)
    for i in idxs:
        if _values_match(lookup_value, keys.at_flat(i)):
            return vals.at_flat(i)

    return XlError.NA if if_not_found is None else if_not_found
