"""Sentinel-returning lookup wrappers for the FormulaEvaluator path."""

from __future__ import annotations

from typing import cast

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.lookup_funcs import (
    hlookup_cells,
    index_cells,
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


def xl_lookup(
    lookup_value: CellValue,
    lookup_vector_or_array: object,
    result_vector: object = None,
) -> CellValue:
    return cast(
        CellValue,
        lookup_cells(lookup_value, lookup_vector_or_array, result_vector),
    )


def xl_index(array: object, row_num: CellValue = None, col_num: CellValue = None) -> CellValue:
    """INDEX over a lazy `Range`, nested list, or materialized grid."""
    return cast(CellValue, index_cells(array, row_num, col_num))


def xl_match(
    lookup_value: CellValue, lookup_array: object, match_type: CellValue = 1
) -> int | XlError:
    return match_cells(lookup_value, lookup_array, match_type)


def xl_vlookup(
    lookup_value: CellValue,
    table_array: object,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    return cast(
        CellValue,
        vlookup_cells(lookup_value, table_array, col_index_num, range_lookup),
    )


def xl_hlookup(
    lookup_value: CellValue,
    table_array: object,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    return cast(
        CellValue,
        hlookup_cells(lookup_value, table_array, row_index_num, range_lookup),
    )


def xl_xlookup(
    lookup_value: CellValue,
    lookup_array: object,
    return_array: object,
    if_not_found: CellValue = None,
    match_mode: CellValue = 0,
    search_mode: CellValue = 1,
) -> CellValue:
    return cast(
        CellValue,
        xlookup_cells(
            lookup_value,
            lookup_array,
            return_array,
            if_not_found,
            match_mode,
            search_mode,
        ),
    )
