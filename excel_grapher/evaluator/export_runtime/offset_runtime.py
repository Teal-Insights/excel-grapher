from __future__ import annotations

import fastpyxl.utils.cell
import numpy as np

from excel_grapher.core.addressing import index_excel_range, offset_range

from .cache import EvalContext, xl_cell
from .core import CellValue, ExcelRange, XlError, to_number


def _quote_sheet_if_needed(sheet: str) -> str:
    if " " in sheet or "-" in sheet or "'" in sheet:
        return f"'{sheet}'"
    return sheet


def _format_address(sheet: str, row: int, col: int) -> str:
    sheet_name = _quote_sheet_if_needed(sheet)
    col_letter = fastpyxl.utils.cell.get_column_letter(col)
    return f"{sheet_name}!{col_letter}{row}"


def xl_offset_ref(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
) -> ExcelRange | XlError:
    if isinstance(ref, ExcelRange):
        base_range = ref
    else:
        match ref:
            case (sheet, base_row, base_col):
                base_range = ExcelRange(
                    sheet=sheet,
                    start_row=base_row,
                    start_col=base_col,
                    end_row=base_row,
                    end_col=base_col,
                )
            case (sheet, base_row, base_col, base_end_row, base_end_col):
                base_range = ExcelRange(
                    sheet=sheet,
                    start_row=base_row,
                    start_col=base_col,
                    end_row=base_end_row,
                    end_col=base_end_col,
                )
            case _:
                return XlError.VALUE

    # Use unbounded workbook limits here; runtime callers are responsible for
    # ensuring they do not construct invalid coordinates.
    class _UnboundedSheet:
        sheet = base_range.sheet
        min_row = 1
        min_col = 1
        max_row = 1_000_000_000
        max_col = 1_000_000_000

    return offset_range(
        base_range,
        rows,
        cols,
        height,
        width,
        bounds=_UnboundedSheet(),
    )


def xl_index_ref(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
    row_num: CellValue | None,
    col_num: CellValue | None,
) -> ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int] | XlError:
    """INDEX semantics that return a reference suitable for OFFSET."""
    if isinstance(ref, ExcelRange):
        base = ref
    else:
        match ref:
            case (sheet, r1, c1):
                base = ExcelRange(sheet=sheet, start_row=r1, start_col=c1, end_row=r1, end_col=c1)
            case (sheet, r1, c1, r2, c2):
                base = ExcelRange(sheet=sheet, start_row=r1, start_col=c1, end_row=r2, end_col=c2)
            case _:
                return XlError.VALUE

    out = index_excel_range(base, row_num, col_num)
    if isinstance(out, XlError):
        return out
    if out.start_row == out.end_row and out.start_col == out.end_col:
        return (out.sheet, out.start_row, out.start_col)
    return (out.sheet, out.start_row, out.start_col, out.end_row, out.end_col)


def xl_offset(
    ctx: EvalContext,
    ref_info: tuple[str, int, int] | tuple[str, int, int, int, int],
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
) -> CellValue:
    rr = to_number(rows)
    if isinstance(rr, XlError):
        return rr
    cc = to_number(cols)
    if isinstance(cc, XlError):
        return cc

    match ref_info:
        case (sheet, base_row, base_col):
            base_end_row, base_end_col = base_row, base_col
        case (sheet, base_row, base_col, base_end_row, base_end_col):
            pass

    base_h = int(base_end_row - base_row + 1)
    base_w = int(base_end_col - base_col + 1)

    if height is None:
        h = base_h
    else:
        hh = to_number(height)
        if isinstance(hh, XlError):
            return hh
        h = int(hh)

    if width is None:
        w = base_w
    else:
        ww = to_number(width)
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
        return xl_cell(ctx, addr)

    result: list[list[CellValue]] = []
    for r in range(target_row, target_row + h):
        row_values: list[CellValue] = []
        for c in range(target_col, target_col + w):
            addr = _format_address(sheet, r, c)
            row_values.append(xl_cell(ctx, addr))
        result.append(row_values)
    return np.array(result, dtype=object)
