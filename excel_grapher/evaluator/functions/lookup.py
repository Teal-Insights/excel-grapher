from __future__ import annotations

import numpy as np

from ..export_runtime.lookup import (
    xl__xlfn_xlookup as _rt_xl__xlfn_xlookup,
)
from ..export_runtime.lookup import (
    xl_index as _rt_xl_index,
)
from ..export_runtime.lookup import (
    xl_lookup as _rt_xl_lookup,
)
from ..export_runtime.lookup import (
    xl_match as _rt_xl_match,
)
from ..helpers import excel_casefold, to_native, to_number
from ..types import CellValue, XlError
from . import register


@register("INDEX")
def xl_index(
    array: np.ndarray, row_num: CellValue, col_num: CellValue = None
) -> CellValue:
    """Return a value from a specific position in an array."""
    return _rt_xl_index(array, row_num, col_num)


@register("MATCH")
def xl_match(
    lookup_value: CellValue,
    lookup_array: CellValue,
    match_type: CellValue = 1,
) -> int | XlError:
    """Find the position of a value in a range."""
    return _rt_xl_match(lookup_value, lookup_array, match_type)


def _values_match(a: CellValue, b: CellValue) -> bool:
    """Check if two values match (case-insensitive for strings)."""
    if isinstance(a, str) and isinstance(b, str):
        return excel_casefold(a) == excel_casefold(b)
    # For numbers, compare numerically
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return an == bn
    # Fallback to direct comparison
    return a == b


def _compare_values(a: CellValue, b: CellValue) -> int:
    """Compare two values. Returns <0 if a<b, 0 if a==b, >0 if a>b."""
    # For numbers, compare numerically
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        if an < bn:
            return -1
        if an > bn:
            return 1
        return 0
    # For strings, compare case-insensitively
    if isinstance(a, str) and isinstance(b, str):
        af = excel_casefold(a)
        bf = excel_casefold(b)
        if af < bf:
            return -1
        if af > bf:
            return 1
        return 0
    # Mixed types - numbers < strings
    return 0


@register("LOOKUP")
def xl_lookup(
    lookup_value: CellValue,
    lookup_vector_or_array: np.ndarray,
    result_vector: np.ndarray | None = None,
) -> CellValue:
    return _rt_xl_lookup(lookup_value, lookup_vector_or_array, result_vector)


@register("VLOOKUP")
def xl_vlookup(
    lookup_value: CellValue,
    table_array: np.ndarray,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    """Search for a value in the first column and return a value in the same row from another column."""
    # Convert col_index_num
    cn = to_number(col_index_num)
    if isinstance(cn, XlError):
        return cn
    col_index = int(cn)

    if col_index < 1:
        return XlError.VALUE

    rows, cols = table_array.shape
    if col_index > cols:
        return XlError.REF

    # Determine match type
    exact_match = not range_lookup

    # Search first column
    first_col = table_array[:, 0]

    if exact_match:
        # Exact match
        for i, val in enumerate(first_col):
            if _values_match(lookup_value, val):
                return to_native(table_array[i, col_index - 1])
        return XlError.NA
    else:
        # Approximate match - find largest value <= lookup_value
        last_match_idx = None
        for i, val in enumerate(first_col):
            if _compare_values(val, lookup_value) <= 0:
                last_match_idx = i
            else:
                break
        if last_match_idx is None:
            return XlError.NA
        return to_native(table_array[last_match_idx, col_index - 1])


@register("HLOOKUP")
def xl_hlookup(
    lookup_value: CellValue,
    table_array: np.ndarray,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    """Search for a value in the first row and return a value in the same column from another row."""
    # Convert row_index_num
    rn = to_number(row_index_num)
    if isinstance(rn, XlError):
        return rn
    row_index = int(rn)

    if row_index < 1:
        return XlError.VALUE

    rows, cols = table_array.shape
    if row_index > rows:
        return XlError.REF

    # Determine match type
    exact_match = not range_lookup

    # Search first row
    first_row = table_array[0, :]

    if exact_match:
        # Exact match
        for i, val in enumerate(first_row):
            if _values_match(lookup_value, val):
                return to_native(table_array[row_index - 1, i])
        return XlError.NA
    else:
        # Approximate match - find largest value <= lookup_value
        last_match_idx = None
        for i, val in enumerate(first_row):
            if _compare_values(val, lookup_value) <= 0:
                last_match_idx = i
            else:
                break
        if last_match_idx is None:
            return XlError.NA
        return to_native(table_array[row_index - 1, last_match_idx])


@register("_XLFN.XLOOKUP")
@register("XLOOKUP")
def xl__xlfn_xlookup(
    lookup_value: CellValue,
    lookup_array: np.ndarray,
    return_array: np.ndarray,
    if_not_found: CellValue = None,
    match_mode: CellValue = 0,
    search_mode: CellValue = 1,
) -> CellValue:
    return _rt_xl__xlfn_xlookup(
        lookup_value,
        lookup_array,
        return_array,
        if_not_found,
        match_mode,
        search_mode,
    )
