"""Sentinel-returning lookup wrappers for the FormulaEvaluator path."""

from __future__ import annotations

import numpy as np

from excel_grapher.core import CellValue, XlError, to_native, to_number
from excel_grapher.core.lookup_funcs import (
    hlookup_cells,
    lookup_cells,
    match_cells,
    vlookup_cells,
    xlookup_cells,
)

__all__ = [
    "xl_hlookup",
    "xl_index",
    "xl_lookup",
    "xl_match",
    "xl_vlookup",
    "xl_xlookup",
]


def _array_arg(value: object) -> object:
    """Materialize numpy arrays to nested lists for shared Grid consumers."""
    if isinstance(value, np.ndarray):
        return value.tolist()
    return value


def xl_lookup(
    lookup_value: CellValue,
    lookup_vector_or_array: object,
    result_vector: object = None,
) -> CellValue:
    return lookup_cells(lookup_value, _array_arg(lookup_vector_or_array), _array_arg(result_vector))


def xl_index(array: np.ndarray, row_num: CellValue, col_num: CellValue = None) -> CellValue:
    """INDEX over a materialized ndarray (legacy path; evaluator uses geometry)."""
    if not isinstance(array, np.ndarray):
        return XlError.VALUE
    nrows, ncols = array.shape
    row_omitted = row_num is None
    col_omitted = col_num is None

    if row_omitted and col_omitted:
        if nrows == 1 and ncols == 1:
            return to_native(array[0, 0])
        if nrows == 1:
            return to_native(array[0, ncols - 1])
        if ncols == 1:
            return to_native(array[nrows - 1, 0])
        return XlError.VALUE

    if row_omitted:
        cn = to_number(col_num)
        if isinstance(cn, XlError):
            return cn
        col = int(cn)
        if col < 1 or col > ncols:
            return XlError.REF
        if nrows == 1:
            return to_native(array[0, col - 1])
        return array[:, col - 1 : col]

    rn = to_number(row_num)
    if isinstance(rn, XlError):
        return rn
    row = int(rn)

    if col_omitted:
        if nrows == 1:
            if row < 1 or row > ncols:
                return XlError.REF
            return to_native(array[0, row - 1])
        if ncols == 1:
            if row < 1 or row > nrows:
                return XlError.REF
            return to_native(array[row - 1, 0])
        if row < 1 or row > nrows:
            return XlError.REF
        return array[row - 1 : row, :]

    cn = to_number(col_num)
    if isinstance(cn, XlError):
        return cn
    col = int(cn)
    if nrows == 1:
        if row < 1 or row > ncols:
            return XlError.REF
        return to_native(array[0, row - 1])
    if ncols == 1:
        if row < 1 or row > nrows:
            return XlError.REF
        return to_native(array[row - 1, 0])
    if row < 1 or row > nrows:
        return XlError.REF
    if col < 1 or col > ncols:
        return XlError.REF
    return to_native(array[row - 1, col - 1])


def xl_match(
    lookup_value: CellValue, lookup_array: object, match_type: CellValue = 1
) -> int | XlError:
    return match_cells(lookup_value, _array_arg(lookup_array), match_type)


def xl_vlookup(
    lookup_value: CellValue,
    table_array: object,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    return vlookup_cells(lookup_value, _array_arg(table_array), col_index_num, range_lookup)


def xl_hlookup(
    lookup_value: CellValue,
    table_array: object,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    return hlookup_cells(lookup_value, _array_arg(table_array), row_index_num, range_lookup)


def xl_xlookup(
    lookup_value: CellValue,
    lookup_array: object,
    return_array: object,
    if_not_found: CellValue = None,
    match_mode: CellValue = 0,
    search_mode: CellValue = 1,
) -> CellValue:
    return xlookup_cells(
        lookup_value,
        _array_arg(lookup_array),
        _array_arg(return_array),
        if_not_found,
        match_mode,
        search_mode,
    )
