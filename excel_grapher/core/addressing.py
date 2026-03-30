from __future__ import annotations

from typing import Protocol

import fastpyxl.utils.cell

from . import CellValue, ExcelRange, XlError, to_number


class WorkbookBoundsProtocol(Protocol):
    """Minimal protocol for workbook/sheet bounds used by addressing helpers."""

    sheet: str
    min_row: int
    max_row: int
    min_col: int
    max_col: int


def _in_bounds(rng: ExcelRange, bounds: WorkbookBoundsProtocol) -> bool:
    if rng.sheet != bounds.sheet:
        return False
    return (
        bounds.min_row <= rng.start_row <= rng.end_row <= bounds.max_row
        and bounds.min_col <= rng.start_col <= rng.end_col <= bounds.max_col
    )


def offset_range(
    base: ExcelRange,
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
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

    # Optional sheet qualifier.
    if "!" in raw:
        sheet_text, addr_text = raw.split("!", 1)
        sheet = sheet_text or bounds.sheet
    else:
        sheet = bounds.sheet
        addr_text = raw

    try:
        if ":" in addr_text:
            start_ref, end_ref = addr_text.split(":", 1)
            start_col, start_row = fastpyxl.utils.cell.coordinate_to_tuple(start_ref)
            end_col, end_row = fastpyxl.utils.cell.coordinate_to_tuple(end_ref)
        else:
            col, row = fastpyxl.utils.cell.coordinate_to_tuple(addr_text)
            start_col = end_col = col
            start_row = end_row = row
    except Exception:
        return XlError.NAME

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
