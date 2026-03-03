from __future__ import annotations

import numpy as np

from .core import CellValue, XlError, excel_casefold, to_native, to_number


def _values_match(a: CellValue, b: CellValue) -> bool:
    if isinstance(a, str) and isinstance(b, str):
        return excel_casefold(a) == excel_casefold(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return an == bn
    return a == b


def _compare_values(a: CellValue, b: CellValue) -> int:
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return -1 if an < bn else 1 if an > bn else 0
    if isinstance(a, str) and isinstance(b, str):
        af = excel_casefold(a)
        bf = excel_casefold(b)
        return -1 if af < bf else 1 if af > bf else 0
    return 0


def _as_vector(array: np.ndarray) -> tuple[np.ndarray, bool]:
    if array.ndim == 1:
        return array, True
    rows, cols = array.shape
    if rows == 1:
        return array[0, :], True
    if cols == 1:
        return array[:, 0], True
    return array, False


def xl_lookup(
    lookup_value: CellValue,
    lookup_vector_or_array: np.ndarray,
    result_vector: np.ndarray | None = None,
) -> CellValue:
    if not isinstance(lookup_vector_or_array, np.ndarray):
        return XlError.VALUE
    if result_vector is not None and not isinstance(result_vector, np.ndarray):
        return XlError.VALUE

    if result_vector is None:
        lookup_vec, is_vector = _as_vector(lookup_vector_or_array)
        if is_vector:
            result_vec = lookup_vec
        else:
            rows, cols = lookup_vector_or_array.shape
            if rows >= cols:
                lookup_vec = lookup_vector_or_array[:, 0]
                result_vec = lookup_vector_or_array[:, -1]
            else:
                lookup_vec = lookup_vector_or_array[0, :]
                result_vec = lookup_vector_or_array[-1, :]
    else:
        lookup_vec, is_vector = _as_vector(lookup_vector_or_array)
        result_vec, is_result_vector = _as_vector(result_vector)
        if not is_vector or not is_result_vector:
            return XlError.NA
        if lookup_vec.shape[0] != result_vec.shape[0]:
            return XlError.NA

    last_match_idx = None
    for i, val in enumerate(lookup_vec):
        if _compare_values(to_native(val), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return to_native(result_vec[last_match_idx])


def xl_index(array: np.ndarray, row_num: CellValue, col_num: CellValue = None) -> CellValue:
    if not isinstance(array, np.ndarray):
        return XlError.VALUE
    rn = to_number(row_num)
    if isinstance(rn, XlError):
        return rn
    row = int(rn)
    if col_num is None:
        col = 1
    else:
        cn = to_number(col_num)
        if isinstance(cn, XlError):
            return cn
        col = int(cn)
    rows, cols = array.shape
    if rows == 1:
        if row < 1 or row > cols:
            return XlError.REF
        return to_native(array[0, row - 1])
    if cols == 1:
        if row < 1 or row > rows:
            return XlError.REF
        return to_native(array[row - 1, 0])
    if row < 1 or row > rows:
        return XlError.REF
    if col < 1 or col > cols:
        return XlError.REF
    return to_native(array[row - 1, col - 1])


def xl_match(
    lookup_value: CellValue, lookup_array: CellValue, match_type: CellValue = 1
) -> int | XlError:
    mt = to_number(match_type)
    if isinstance(mt, XlError):
        return mt
    match_type_int = int(mt)
    if isinstance(lookup_array, XlError):
        return lookup_array
    if isinstance(lookup_array, np.ndarray):
        flat = np.ravel(lookup_array)
    elif isinstance(lookup_array, (list, tuple)):
        flat = np.ravel(np.array(lookup_array, dtype=object))
    else:
        flat = np.array([lookup_array], dtype=object)
    if match_type_int == 0:
        for i, val in enumerate(flat):
            if _values_match(lookup_value, val):
                return i + 1
        return XlError.NA
    if match_type_int == 1:
        last_match = None
        for i, val in enumerate(flat):
            if _compare_values(val, lookup_value) <= 0:
                last_match = i + 1
            else:
                break
        return XlError.NA if last_match is None else last_match
    if match_type_int == -1:
        last_match = None
        for i, val in enumerate(flat):
            if _compare_values(val, lookup_value) >= 0:
                last_match = i + 1
            else:
                break
        return XlError.NA if last_match is None else last_match
    return XlError.VALUE


def xl_vlookup(
    lookup_value: CellValue,
    table_array: np.ndarray,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    cn = to_number(col_index_num)
    if isinstance(cn, XlError):
        return cn
    col_index = int(cn)
    if col_index < 1:
        return XlError.VALUE
    rows, cols = table_array.shape
    if col_index > cols:
        return XlError.REF
    exact_match = not bool(range_lookup)
    first_col = table_array[:, 0]
    if exact_match:
        for i, val in enumerate(first_col):
            if _values_match(lookup_value, val):
                return to_native(table_array[i, col_index - 1])
        return XlError.NA
    last_match_idx = None
    for i, val in enumerate(first_col):
        if _compare_values(val, lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return to_native(table_array[last_match_idx, col_index - 1])


def xl_hlookup(
    lookup_value: CellValue,
    table_array: np.ndarray,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    rn = to_number(row_index_num)
    if isinstance(rn, XlError):
        return rn
    row_index = int(rn)
    if row_index < 1:
        return XlError.VALUE
    rows, cols = table_array.shape
    if row_index > rows:
        return XlError.REF
    exact_match = not bool(range_lookup)
    first_row = table_array[0, :]
    if exact_match:
        for i, val in enumerate(first_row):
            if _values_match(lookup_value, val):
                return to_native(table_array[row_index - 1, i])
        return XlError.NA
    last_match_idx = None
    for i, val in enumerate(first_row):
        if _compare_values(val, lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return to_native(table_array[row_index - 1, last_match_idx])


def xl__xlfn_xlookup(
    lookup_value: CellValue,
    lookup_array: np.ndarray,
    return_array: np.ndarray,
    if_not_found: CellValue = None,
    match_mode: CellValue = 0,
    search_mode: CellValue = 1,
) -> CellValue:
    """Excel XLOOKUP (including `_xlfn.XLOOKUP`).

    This implementation supports:
    - exact match (match_mode=0)
    - search first-to-last (search_mode=1) and last-to-first (search_mode=-1)
    """
    mm = to_number(match_mode)
    if isinstance(mm, XlError):
        return mm
    sm = to_number(search_mode)
    if isinstance(sm, XlError):
        return sm

    mm_i = int(mm)
    sm_i = int(sm)

    if mm_i != 0:
        return XlError.VALUE
    if sm_i not in (1, -1):
        return XlError.VALUE

    keys = np.ravel(lookup_array)
    vals = np.ravel(return_array)
    if keys.shape[0] != vals.shape[0]:
        return XlError.VALUE

    idxs = range(keys.shape[0]) if sm_i == 1 else range(keys.shape[0] - 1, -1, -1)
    for i in idxs:
        if _values_match(lookup_value, to_native(keys[i])):
            return to_native(vals[i])

    return XlError.NA if if_not_found is None else if_not_found

