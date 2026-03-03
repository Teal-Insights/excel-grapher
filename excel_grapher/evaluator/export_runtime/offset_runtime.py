from __future__ import annotations

import numpy as np
import openpyxl.utils.cell

from .cache import EvalContext, xl_cell
from .core import CellValue, ExcelRange, XlError, to_number


def _quote_sheet_if_needed(sheet: str) -> str:
    if " " in sheet or "-" in sheet or "'" in sheet:
        return f"'{sheet}'"
    return sheet


def _format_address(sheet: str, row: int, col: int) -> str:
    sheet_name = _quote_sheet_if_needed(sheet)
    col_letter = openpyxl.utils.cell.get_column_letter(col)
    return f"{sheet_name}!{col_letter}{row}"


def xl_offset_ref(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
) -> ExcelRange | XlError:
    if isinstance(ref, ExcelRange):
        sheet = ref.sheet
        base_row = ref.start_row
        base_col = ref.start_col
        base_end_row = ref.end_row
        base_end_col = ref.end_col
    else:
        match ref:
            case (sheet, base_row, base_col):
                base_end_row, base_end_col = base_row, base_col
            case (sheet, base_row, base_col, base_end_row, base_end_col):
                pass
            case _:
                return XlError.VALUE

    rr = to_number(rows)
    if isinstance(rr, XlError):
        return rr
    cc = to_number(cols)
    if isinstance(cc, XlError):
        return cc

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

    return ExcelRange(
        sheet=sheet,
        start_row=target_row,
        start_col=target_col,
        end_row=target_row + h - 1,
        end_col=target_col + w - 1,
    )


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

