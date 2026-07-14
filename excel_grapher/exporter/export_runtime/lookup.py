"""Lookup and reference functions over lazy ranges for exported code."""

from __future__ import annotations

from typing import cast

from excel_grapher.core.lookup_funcs import (
    hlookup_cells,
    index_cells,
    lookup_cells,
    match_cells,
    vlookup_cells,
    xlookup_cells,
)
from excel_grapher.core.types import XlError, XlErrorException

from .values import CellValue

__all__ = [
    "xl_hlookup",
    "xl_index",
    "xl_lookup",
    "xl_match",
    "xl_vlookup",
    "xl_xlookup",
]


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


def xl_index(array: CellValue, row_num: CellValue = None, col_num: CellValue = None) -> CellValue:
    return _raise_if_error(index_cells(array, row_num, col_num))


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
