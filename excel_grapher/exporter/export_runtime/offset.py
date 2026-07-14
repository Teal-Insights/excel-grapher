"""OFFSET and workbook-range materialization over lazy ranges for exported code."""

from __future__ import annotations

from typing import cast

import fastpyxl.utils.cell

from excel_grapher.core import XlError, to_number
from excel_grapher.core.address_keys import format_cell_key
from excel_grapher.core.addressing import index_excel_range, offset_range
from excel_grapher.core.types import XlErrorException
from excel_grapher.runtime.cache import EvalContext, _parse_range_address, xl_cell

from .ranges import Range
from .values import CellValue, ExcelRange, Scalar, as_scalar

__all__ = ["xl_index_ref", "xl_offset", "xl_offset_ref", "xl_range", "xl_range_rows"]


def _format_address(sheet: str, row: int, col: int) -> str:
    return format_cell_key(sheet, fastpyxl.utils.cell.get_column_letter(col), row)


def _number_or_raise(value: CellValue) -> float:
    """Coerce a scalar argument to a number, raising on Excel coercion errors."""
    number = to_number(as_scalar(value))
    if isinstance(number, XlError):
        raise XlErrorException(number)
    return number


def _ctx_range(ctx: EvalContext, sheet: str, r1: int, c1: int, r2: int, c2: int) -> Range:
    # Leave the resolver unannotated: embed strips `excel_grapher.core` imports, so
    # aliases like `CellValue as CoreCellValue` never appear in generated runtime.py.
    def resolve(address: str):
        return xl_cell(ctx, address)

    return Range(sheet, r1, c1, r2, c2, resolve)


def _range_from_ref_info(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
) -> ExcelRange:
    """Normalize generated reference metadata into an `ExcelRange`."""
    if isinstance(ref, ExcelRange):
        return ref
    match ref:
        case (sheet, base_row, base_col):
            return ExcelRange(
                sheet=sheet,
                start_row=base_row,
                start_col=base_col,
                end_row=base_row,
                end_col=base_col,
            )
        case (sheet, base_row, base_col, base_end_row, base_end_col):
            return ExcelRange(
                sheet=sheet,
                start_row=base_row,
                start_col=base_col,
                end_row=base_end_row,
                end_col=base_end_col,
            )
        case _:
            raise XlErrorException(XlError.VALUE)


def _as_addressing_scalar(value: CellValue | None) -> Scalar | None:
    """Collapse export-runtime values to scalars for shared addressing helpers."""
    if value is None:
        return None
    return as_scalar(value)


def _export_range_from_geometry(
    sheet: str, start_row: int, start_col: int, end_row: int, end_col: int
) -> ExcelRange:
    """Build an export-runtime `ExcelRange` from absolute coordinates."""
    return ExcelRange(
        sheet=sheet,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
    )


def xl_index_ref(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
    row_num: CellValue | None,
    col_num: CellValue | None,
) -> tuple[str, int, int] | tuple[str, int, int, int, int]:
    """Return INDEX reference metadata, raising on Excel reference errors."""
    out = index_excel_range(
        _range_from_ref_info(ref),
        _as_addressing_scalar(row_num),
        _as_addressing_scalar(col_num),
    )
    if isinstance(out, XlError):
        raise XlErrorException(out)
    if out.start_row == out.end_row and out.start_col == out.end_col:
        return (out.sheet, out.start_row, out.start_col)
    return (out.sheet, out.start_row, out.start_col, out.end_row, out.end_col)


def xl_offset_ref(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
) -> ExcelRange:
    """Return OFFSET reference metadata, raising on Excel reference errors."""
    base_range = _range_from_ref_info(ref)

    class _UnboundedSheet:
        sheet = base_range.sheet
        min_row = 1
        min_col = 1
        max_row = 1_000_000_000
        max_col = 1_000_000_000

    out = offset_range(
        base_range,
        as_scalar(rows),
        as_scalar(cols),
        _as_addressing_scalar(height),
        _as_addressing_scalar(width),
        bounds=_UnboundedSheet(),
    )
    if isinstance(out, XlError):
        raise XlErrorException(out)
    return _export_range_from_geometry(
        out.sheet, out.start_row, out.start_col, out.end_row, out.end_col
    )


def xl_offset(
    ctx: EvalContext,
    ref_info: tuple[str, int, int] | tuple[str, int, int, int, int] | XlError,
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
) -> CellValue:
    rr = _number_or_raise(rows)
    cc = _number_or_raise(cols)

    if isinstance(ref_info, XlError):
        raise XlErrorException(ref_info)

    match ref_info:
        case (sheet, base_row, base_col):
            base_end_row, base_end_col = base_row, base_col
        case (sheet, base_row, base_col, base_end_row, base_end_col):
            pass
        case _:
            raise XlErrorException(XlError.VALUE)

    base_h = int(base_end_row - base_row + 1)
    base_w = int(base_end_col - base_col + 1)

    h = base_h if height is None else int(_number_or_raise(height))
    w = base_w if width is None else int(_number_or_raise(width))

    target_row = int(base_row + int(rr))
    target_col = int(base_col + int(cc))

    if target_row < 1 or target_col < 1:
        raise XlErrorException(XlError.REF)
    if h <= 0 or w <= 0:
        raise XlErrorException(XlError.VALUE)

    if h == 1 and w == 1:
        addr = _format_address(sheet, target_row, target_col)
        # Core `CellValue` includes core `ExcelRange`; export `CellValue` uses the
        # export-runtime geometry type. Scalar results are always assignable.
        return cast("CellValue", xl_cell(ctx, addr))

    return _ctx_range(ctx, sheet, target_row, target_col, target_row + h - 1, target_col + w - 1)


def xl_range(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a sheet-qualified range address into a lazy `Range` value."""
    parsed = _parse_range_address(address)
    if isinstance(parsed, XlError):
        raise XlErrorException(parsed)
    sheet, start_cell, end_cell = parsed
    try:
        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell)
        start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)
    except ValueError:
        raise XlErrorException(XlError.VALUE) from None

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
