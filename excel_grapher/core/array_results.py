"""Helpers for top-level array formula results (issue #284).

Top-level multi-cell results stay as ``ndarray`` values on the anchor cell.
Range operands project spilled scalars from cached anchors. Occupied spill slots
in the graph return ``XlError.SPILL`` at the anchor.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import cast

import fastpyxl.utils.cell
import numpy as np

from excel_grapher.core.address_keys import format_key, parse_address
from excel_grapher.core.types import CellValue, XlError


def is_array_result(value: CellValue) -> bool:
    """Return whether ``value`` is a multi-cell array formula result."""
    return isinstance(value, np.ndarray) and value.size > 1


def spill_footprint_addresses(anchor_address: str, shape: tuple[int, ...]) -> list[str]:
    """Return sheet-qualified addresses covered by an anchor array spill."""
    sheet, anchor_row, anchor_col = _cell_row_col(anchor_address)
    if len(shape) != 2:
        return [anchor_address]
    rows, cols = int(shape[0]), int(shape[1])
    if rows == 1 and cols == 1:
        return [anchor_address]
    addresses: list[str] = []
    for row_delta in range(rows):
        for col_delta in range(cols):
            row = anchor_row + row_delta
            col = anchor_col + col_delta
            col_text = fastpyxl.utils.cell.get_column_letter(col)
            addresses.append(format_key(sheet, f"{col_text}{row}"))
    return addresses


def spill_blocked_at_anchor(
    anchor_address: str,
    shape: tuple[int, ...],
    is_occupied: Callable[[str], bool],
) -> bool:
    """Return whether any non-anchor spill slot is occupied."""
    for address in spill_footprint_addresses(anchor_address, shape):
        if address == anchor_address:
            continue
        if is_occupied(address):
            return True
    return False


def finalize_top_level_array_result(
    anchor_address: str,
    result: CellValue,
    *,
    is_occupied: Callable[[str], bool],
) -> CellValue:
    """Apply spill blocking to a top-level formula result before caching."""
    if not is_array_result(result):
        return result
    array = cast(np.ndarray, result)
    if spill_blocked_at_anchor(anchor_address, array.shape, is_occupied):
        return XlError.SPILL
    return result


def spill_offsets(
    anchor_row: int,
    anchor_col: int,
    target_row: int,
    target_col: int,
    shape: tuple[int, ...],
) -> tuple[int, int] | None:
    """Return array indices for ``target`` within an anchor spill footprint."""
    if len(shape) != 2:
        return None
    rows = int(shape[0])
    cols = int(shape[1])
    row_delta = target_row - anchor_row
    col_delta = target_col - anchor_col
    if rows > 1 and cols == 1:
        if col_delta != 0 or row_delta < 0 or row_delta >= rows:
            return None
        return (row_delta, 0)
    if cols > 1 and rows == 1:
        if row_delta != 0 or col_delta < 0 or col_delta >= cols:
            return None
        return (0, col_delta)
    if rows > 1 and cols > 1:
        if row_delta < 0 or col_delta < 0 or row_delta >= rows or col_delta >= cols:
            return None
        return (row_delta, col_delta)
    if rows == 1 and cols == 1:
        if row_delta == 0 and col_delta == 0:
            return (0, 0)
        return None
    return None


def _cell_row_col(address: str) -> tuple[str, int, int]:
    sheet, coord = parse_address(address)
    coord = coord.replace("$", "")
    col_text, row = fastpyxl.utils.cell.coordinate_from_string(coord)
    col = int(fastpyxl.utils.cell.column_index_from_string(col_text))
    return sheet, row, col


def spill_scalar_value(
    target_address: str,
    anchor_address: str,
    array: np.ndarray,
) -> CellValue | None:
    """Return the spilled scalar at ``target_address`` from a cached anchor array."""
    target_sheet, target_row, target_col = _cell_row_col(target_address)
    anchor_sheet, anchor_row, anchor_col = _cell_row_col(anchor_address)
    if target_sheet != anchor_sheet:
        return None
    offsets = spill_offsets(anchor_row, anchor_col, target_row, target_col, array.shape)
    if offsets is None:
        return None
    row_index, col_index = offsets
    return cast(CellValue, array[row_index, col_index])


def read_spill_scalar(
    target_address: str,
    cache: Mapping[str, CellValue],
) -> CellValue | None:
    """Read a spill-slot scalar from cached array anchors when ``target`` has no cell."""
    for anchor_address, value in cache.items():
        if not is_array_result(value):
            continue
        array = cast(np.ndarray, value)
        scalar = spill_scalar_value(target_address, anchor_address, array)
        if scalar is not None:
            return scalar
    return None


def scalar_for_range_member(address: str, value: CellValue) -> CellValue:
    """Project a stored cell value to the scalar used inside range operands."""
    if not is_array_result(value):
        return value
    array = cast(np.ndarray, value)
    scalar = spill_scalar_value(address, address, array)
    if scalar is None:
        return value
    return scalar


def array_values_equal(a: object, b: object) -> bool:
    """Compare two values, including element-wise for matching ``ndarray`` shapes."""
    if isinstance(a, np.ndarray) and isinstance(b, np.ndarray):
        left = cast(np.ndarray, a)
        right = cast(np.ndarray, b)
        if left.shape != right.shape:
            return False
        for index in range(int(left.size)):
            if not _scalar_values_equal(left.flat[index], right.flat[index]):
                return False
        return True
    return _scalar_values_equal(a, b)


def _scalar_values_equal(a: object, b: object) -> bool:
    if a is b:
        return True
    if isinstance(a, XlError) or isinstance(b, XlError):
        return a == b
    if isinstance(a, (bool, np.bool_)) and isinstance(b, (bool, np.bool_)):
        return bool(a) == bool(b)
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return float(a) == float(b)
    return a == b
