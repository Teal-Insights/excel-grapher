"""Lookup and reference functions over lazy ranges for exported code."""

from __future__ import annotations

from typing import cast

from excel_grapher.core import XlError, to_number
from excel_grapher.core.lookup_funcs import (
    hlookup_cells,
    lookup_cells,
    match_cells,
    vlookup_cells,
    xlookup_cells,
)
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


def _raise_if_error(value: object) -> CellValue:
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return cast(CellValue, value)


def xl_lookup(
    lookup_value: CellValue,
    lookup_vector_or_array: CellValue,
    result_vector: CellValue = None,
) -> CellValue:
    return _raise_if_error(lookup_cells(lookup_value, lookup_vector_or_array, result_vector))


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
    result = _raise_if_error(match_cells(lookup_value, lookup_array, match_type))
    return cast(int, result)


def xl_vlookup(
    lookup_value: CellValue,
    table_array: CellValue,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    return _raise_if_error(vlookup_cells(lookup_value, table_array, col_index_num, range_lookup))


def xl_hlookup(
    lookup_value: CellValue,
    table_array: CellValue,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    return _raise_if_error(hlookup_cells(lookup_value, table_array, row_index_num, range_lookup))


def xl_xlookup(
    lookup_value: CellValue,
    lookup_array: CellValue,
    return_array: CellValue,
    if_not_found: CellValue = None,
    match_mode: CellValue = 0,
    search_mode: CellValue = 1,
) -> CellValue:
    return _raise_if_error(
        xlookup_cells(
            lookup_value,
            lookup_array,
            return_array,
            if_not_found,
            match_mode,
            search_mode,
        )
    )
