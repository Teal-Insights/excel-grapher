from __future__ import annotations

from typing import Protocol

import fastpyxl.utils.cell

from . import ExcelRange, FormulaValue, XlError, to_number
from .address_keys import parse_address


class ExcelRangeGeometry(Protocol):
    """Minimal rectangular geometry shared by core and export-runtime ranges."""

    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int


def index_excel_range(
    base: ExcelRangeGeometry,
    row_num: FormulaValue | None,
    col_num: FormulaValue | None,
) -> ExcelRange | XlError:
    """Map INDEX(row,col) over *base* to an absolute range (single cell or slice).

    Mirrors `excel_grapher.runtime.lookup.xl_index` geometry
    so OFFSET(INDEX(...), ...) receives a true cell reference.

    `row_num = 0` / `col_num = 0` select the entire opposite axis (Excel's
    whole-column / whole-row form); both zeros return the full base range.
    """
    nrows = base.end_row - base.start_row + 1
    ncols = base.end_col - base.start_col + 1
    row_omitted = row_num is None
    col_omitted = col_num is None

    def abs_cell(r0: int, c0: int) -> ExcelRange:
        r = base.start_row + r0
        c = base.start_col + c0
        return ExcelRange(base.sheet, r, c, r, c)

    def full_base() -> ExcelRange:
        return ExcelRange(base.sheet, base.start_row, base.start_col, base.end_row, base.end_col)

    if row_omitted and col_omitted:
        if nrows == 1 and ncols == 1:
            return abs_cell(0, 0)
        if nrows == 1:
            return abs_cell(0, ncols - 1)
        if ncols == 1:
            return abs_cell(nrows - 1, 0)
        return XlError.VALUE

    if row_omitted:
        cn = to_number(col_num)
        if isinstance(cn, XlError):
            return cn
        col = int(cn)
        if col == 0:
            return full_base()
        if col < 1 or col > ncols:
            return XlError.REF
        if nrows == 1:
            return abs_cell(0, col - 1)
        c0 = base.start_col + col - 1
        return ExcelRange(base.sheet, base.start_row, c0, base.end_row, c0)

    rn = to_number(row_num)
    if isinstance(rn, XlError):
        return rn
    row = int(rn)

    if col_omitted:
        if row == 0:
            return full_base()
        if nrows == 1:
            if row < 1 or row > ncols:
                return XlError.REF
            return abs_cell(0, row - 1)
        if ncols == 1:
            if row < 1 or row > nrows:
                return XlError.REF
            return abs_cell(row - 1, 0)
        if row < 1 or row > nrows:
            return XlError.REF
        r0 = base.start_row + row - 1
        return ExcelRange(base.sheet, r0, base.start_col, r0, base.end_col)

    cn = to_number(col_num)
    if isinstance(cn, XlError):
        return cn
    col = int(cn)

    # Excel: 0 selects the full opposite axis; (0, 0) returns the whole range.
    if row == 0 and col == 0:
        return full_base()
    if row == 0:
        if col < 1 or col > ncols:
            return XlError.REF
        if nrows == 1:
            return abs_cell(0, col - 1)
        c0 = base.start_col + col - 1
        return ExcelRange(base.sheet, base.start_row, c0, base.end_row, c0)
    if col == 0:
        if row < 1 or row > nrows:
            return XlError.REF
        if ncols == 1:
            return abs_cell(row - 1, 0)
        r0 = base.start_row + row - 1
        return ExcelRange(base.sheet, r0, base.start_col, r0, base.end_col)

    if nrows == 1:
        if row < 1 or row > ncols:
            return XlError.REF
        return abs_cell(0, row - 1)
    if ncols == 1:
        if row < 1 or row > nrows:
            return XlError.REF
        return abs_cell(row - 1, 0)
    if row < 1 or row > nrows:
        return XlError.REF
    if col < 1 or col > ncols:
        return XlError.REF
    return abs_cell(row - 1, col - 1)


class WorkbookBoundsProtocol(Protocol):
    """Minimal protocol for workbook/sheet bounds used by addressing helpers."""

    sheet: str
    min_row: int
    max_row: int
    min_col: int
    max_col: int


def split_sheet_qualified_address(address: str) -> tuple[str, str] | None:
    """Split `sheet!coord` into `(sheet_name, coord)`.

    Handles quoted sheet names, including Excel's doubled-single-quote escape
    (`'O''Neil'!A1` -> sheet `O'Neil`).

    Returns `None` when *address* has no sheet qualifier (plain `A1`).
    """
    if "!" not in address:
        return None
    try:
        return parse_address(address)
    except ValueError:
        return None


_split_sheet_qualified_address = split_sheet_qualified_address


def _in_bounds(rng: ExcelRangeGeometry, bounds: WorkbookBoundsProtocol) -> bool:
    if rng.sheet != bounds.sheet:
        return False
    return (
        bounds.min_row <= rng.start_row <= rng.end_row <= bounds.max_row
        and bounds.min_col <= rng.start_col <= rng.end_col <= bounds.max_col
    )


def offset_range(
    base: ExcelRangeGeometry,
    rows: FormulaValue,
    cols: FormulaValue,
    height: FormulaValue | None = None,
    width: FormulaValue | None = None,
    *,
    bounds: WorkbookBoundsProtocol,
) -> ExcelRange | XlError:
    """Compute the Excel OFFSET target range in a representation-agnostic way.

    Semantics are aligned with the canonical runtime implementation used by the
    evaluator and export runtime:
    - Row/column offsets are coerced via to_number and propagate errors.
    - Height/width default to the base range shape when omitted.
    - Non-positive height/width return XlError.VALUE.
    - Targets that land outside the provided bounds return XlError.REF.
    """
    rr = to_number(rows)
    if isinstance(rr, XlError):
        return rr
    cc = to_number(cols)
    if isinstance(cc, XlError):
        return cc

    base_h = int(base.end_row - base.start_row + 1)
    base_w = int(base.end_col - base.start_col + 1)

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

    if h <= 0 or w <= 0:
        return XlError.VALUE

    target_row = int(base.start_row + int(rr))
    target_col = int(base.start_col + int(cc))

    if target_row < 1 or target_col < 1:
        return XlError.REF

    result = ExcelRange(
        sheet=base.sheet,
        start_row=target_row,
        start_col=target_col,
        end_row=target_row + h - 1,
        end_col=target_col + w - 1,
    )
    if not _in_bounds(result, bounds):
        return XlError.REF
    return result


def indirect_text_to_range(
    text: str,
    a1: bool,
    *,
    bounds: WorkbookBoundsProtocol,
) -> ExcelRange | XlError:
    """Interpret an INDIRECT text argument as an ExcelRange.

    This helper currently supports only A1-style references; R1C1 mode is
    treated as unsupported and returns XlError.NAME.
    """
    if not a1:
        # R1C1 mode is currently unsupported in this helper.
        return XlError.NAME

    raw = text.strip()
    if not raw:
        return XlError.NAME

    try:
        if ":" in raw:
            start_text, end_text = raw.split(":", 1)
            parsed_start = split_sheet_qualified_address(start_text)
            if parsed_start is None:
                sheet = bounds.sheet
                start_ref = start_text
            else:
                sheet, start_ref = parsed_start

            parsed_end = split_sheet_qualified_address(end_text)
            if parsed_end is None:
                end_ref = end_text
            else:
                end_sheet, end_ref = parsed_end
                if end_sheet != sheet:
                    return XlError.NAME

            start_col_s, start_row = fastpyxl.utils.cell.coordinate_from_string(start_ref)
            end_col_s, end_row = fastpyxl.utils.cell.coordinate_from_string(end_ref)
            start_col = fastpyxl.utils.cell.column_index_from_string(start_col_s)
            end_col = fastpyxl.utils.cell.column_index_from_string(end_col_s)
        else:
            parsed = split_sheet_qualified_address(raw)
            if parsed is None:
                sheet = bounds.sheet
                ref_text = raw
            else:
                sheet, ref_text = parsed
            col_s, row = fastpyxl.utils.cell.coordinate_from_string(ref_text)
            col = fastpyxl.utils.cell.column_index_from_string(col_s)
            start_col = end_col = col
            start_row = end_row = row
    except Exception:
        return XlError.NAME

    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col > end_col:
        start_col, end_col = end_col, start_col

    rng = ExcelRange(
        sheet=sheet,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
    )
    if not _in_bounds(rng, bounds):
        return XlError.REF
    return rng
