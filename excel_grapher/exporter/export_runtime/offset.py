"""OFFSET and workbook-range materialization over lazy ranges for exported code."""

from __future__ import annotations

from typing import Any, cast

import fastpyxl.utils.cell

from excel_grapher.core import XlError, to_number
from excel_grapher.runtime.cache import EvalContext, _parse_range_address, xl_cell

from .ranges import Range
from .values import CellValue, as_scalar

__all__ = ["xl_offset", "xl_range", "xl_range_rows"]


def _quote_sheet_if_needed(sheet: str) -> str:
    if " " in sheet or "-" in sheet or "'" in sheet:
        return f"'{sheet}'"
    return sheet


def _format_address(sheet: str, row: int, col: int) -> str:
    sheet_name = _quote_sheet_if_needed(sheet)
    col_letter = fastpyxl.utils.cell.get_column_letter(col)
    return f"{sheet_name}!{col_letter}{row}"


def _ctx_range(ctx: EvalContext, sheet: str, r1: int, c1: int, r2: int, c2: int) -> Range:
    def resolve(address: str) -> Any:
        return xl_cell(ctx, address)

    return Range(sheet, r1, c1, r2, c2, resolve)


def xl_offset(
    ctx: EvalContext,
    ref_info: tuple[str, int, int] | tuple[str, int, int, int, int] | XlError,
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
) -> CellValue:
    rr = to_number(as_scalar(rows))
    if isinstance(rr, XlError):
        return rr
    cc = to_number(as_scalar(cols))
    if isinstance(cc, XlError):
        return cc

    if isinstance(ref_info, XlError):
        return ref_info

    match ref_info:
        case (sheet, base_row, base_col):
            base_end_row, base_end_col = base_row, base_col
        case (sheet, base_row, base_col, base_end_row, base_end_col):
            pass
        case _:
            return XlError.VALUE

    base_h = int(base_end_row - base_row + 1)
    base_w = int(base_end_col - base_col + 1)

    if height is None:
        h = base_h
    else:
        hh = to_number(as_scalar(height))
        if isinstance(hh, XlError):
            return hh
        h = int(hh)

    if width is None:
        w = base_w
    else:
        ww = to_number(as_scalar(width))
        if isinstance(ww, XlError):
            return ww
        w = int(ww)

    target_row = int(base_row + int(rr))
    target_col = int(base_col + int(cc))

    if target_row < 1 or target_col < 1:
        return XlError.REF
    if h <= 0 or w <= 0:
        return XlError.VALUE

    if h == 1 and w == 1:
        addr = _format_address(sheet, target_row, target_col)
        return cast("CellValue", xl_cell(ctx, addr))

    return _ctx_range(ctx, sheet, target_row, target_col, target_row + h - 1, target_col + w - 1)


def xl_range(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a sheet-qualified range address into a lazy `Range` value."""
    parsed = _parse_range_address(address)
    if isinstance(parsed, XlError):
        return parsed
    sheet, start_cell, end_cell = parsed
    try:
        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell)
        start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)
    except ValueError:
        return XlError.VALUE

    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col_idx > end_col_idx:
        start_col_idx, end_col_idx = end_col_idx, start_col_idx

    return _ctx_range(ctx, sheet, start_row, start_col_idx, end_row, end_col_idx)


def xl_range_rows(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a sheet-qualified range eagerly into nested row lists.

    Public boundary handler for range targets: results returned from
    `compute_all` are materialized values, not lazy range views.
    """
    rng = xl_range(ctx, address)
    if isinstance(rng, Range):
        return rng.rows_raw()
    return rng
