"""Shared lookup/reference algorithms over lazy `Grid` values.

Implementations return `XlError` sentinels. Runtime wrappers expose them to the
evaluator; export_runtime wrappers raise `XlErrorException` at the boundary.

Evaluator and export value vocabularies differ slightly (export `CellValue`
also includes lazy `Range`); shared traversal accepts opaque objects and
narrows at use sites.
"""

from __future__ import annotations

from typing import cast

from excel_grapher.core.coercions import as_scalar, excel_casefold, to_number
from excel_grapher.core.grid import Grid, Range, Scalar
from excel_grapher.core.types import CellValue, XlError

__all__ = [
    "hlookup_cells",
    "index_cells",
    "lookup_cells",
    "match_cells",
    "vlookup_cells",
    "xlookup_cells",
]


def index_cells(
    array: object,
    row_num: object = None,
    col_num: object = None,
) -> object:
    """Excel INDEX over a lazy grid or nested-list array.

    Returns a scalar cell value, or a row/column slice (`Range` or nested list)
    when only one of `row_num` / `col_num` selects a vector.
    """
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
        col_s = as_scalar(col_num)
        if isinstance(col_s, XlError):
            return col_s
        cn = to_number(cast(CellValue, col_s))
        if isinstance(cn, XlError):
            return cn
        col = int(cn)
        if col < 1 or col > ncols:
            return XlError.REF
        if nrows == 1:
            return grid.at(0, col - 1)
        return grid.col_slice(col - 1)

    row_s = as_scalar(row_num)
    if isinstance(row_s, XlError):
        return row_s
    rn = to_number(cast(CellValue, row_s))
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

    col_s = as_scalar(col_num)
    if isinstance(col_s, XlError):
        return col_s
    cn = to_number(cast(CellValue, col_s))
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


def _as_scalar(value: object) -> Scalar:
    if isinstance(value, (Range, list, tuple)):
        return XlError.VALUE
    if Grid.wrap(value) is not None:
        return XlError.VALUE
    return cast(Scalar, value)


def _values_match(a: object, b: object) -> bool:
    a = _as_scalar(a)
    b = _as_scalar(b)
    if isinstance(a, str) and isinstance(b, str):
        return excel_casefold(a) == excel_casefold(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return an == bn
    return a == b


def _compare_values(a: object, b: object) -> int:
    a = _as_scalar(a)
    b = _as_scalar(b)
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
    if grid.nrows == 1 or grid.ncols == 1:
        return grid
    return None


def _coerce_grid(value: object) -> Grid | XlError | None:
    """Wrap array-like values; return `None` for scalars, errors as-is."""
    if isinstance(value, XlError):
        return value
    return Grid.wrap(value)


def lookup_cells(
    lookup_value: object,
    lookup_vector_or_array: object,
    result_vector: object = None,
) -> Scalar:
    """Excel LOOKUP over a lazy grid or nested-list array."""
    grid = _coerce_grid(lookup_vector_or_array)
    if grid is None:
        return XlError.VALUE
    if isinstance(grid, XlError):
        return grid
    result_grid = _coerce_grid(result_vector) if result_vector is not None else None
    if result_vector is not None and result_grid is None:
        return XlError.VALUE
    if isinstance(result_grid, XlError):
        return result_grid

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


def match_cells(
    lookup_value: object,
    lookup_array: object,
    match_type: object = 1,
) -> int | XlError:
    """Excel MATCH over a lazy grid or nested-list array."""
    mt = to_number(cast(CellValue, match_type))
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


def vlookup_cells(
    lookup_value: object,
    table_array: object,
    col_index_num: object,
    range_lookup: object = True,
) -> Scalar:
    """Excel VLOOKUP over a lazy grid or nested-list array."""
    cn = to_number(cast(CellValue, col_index_num))
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


def hlookup_cells(
    lookup_value: object,
    table_array: object,
    row_index_num: object,
    range_lookup: object = True,
) -> Scalar:
    """Excel HLOOKUP over a lazy grid or nested-list array."""
    rn = to_number(cast(CellValue, row_index_num))
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


def xlookup_cells(
    lookup_value: object,
    lookup_array: object,
    return_array: object,
    if_not_found: object = None,
    match_mode: object = 0,
    search_mode: object = 1,
) -> object:
    """Excel XLOOKUP (exact match; search forward or backward)."""
    mm = to_number(cast(CellValue, match_mode))
    if isinstance(mm, XlError):
        return mm
    sm = to_number(cast(CellValue, search_mode))
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
