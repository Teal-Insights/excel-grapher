"""Parse cell and range references from raw formulas, preserving `$` markers."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

from excel_grapher.grapher.parser import mask_spans, parse_range_refs_with_spans

RefKind = Literal["cell", "range"]


@dataclass(frozen=True, slots=True)
class AbsCellRef:
    """Single cell reference extracted from a raw formula."""

    sheet: str | None
    column: str
    row: int
    is_absolute_col: bool
    is_absolute_row: bool
    span: tuple[int, int] = (0, 0)

    @property
    def kind(self) -> RefKind:
        return "cell"


@dataclass(frozen=True, slots=True)
class AbsRangeRef:
    """Range reference extracted from a raw formula with endpoint absoluteness."""

    sheet: str | None
    start_col: str
    start_row: int
    end_col: str
    end_row: int
    start_abs_col: bool
    start_abs_row: bool
    end_abs_col: bool
    end_abs_row: bool
    span: tuple[int, int] = (0, 0)

    @property
    def kind(self) -> RefKind:
        return "range"


FormulaRef = AbsCellRef | AbsRangeRef


_SHEET_CELL_ABS_RE = re.compile(
    r"(?:'(?P<qs>[^']+)'|(?P<us>[A-Za-z][A-Za-z0-9_]*))!"
    r"(?P<acol>\$)?(?P<col>[A-Z]{1,3})(?P<arow>\$)?(?P<row>\d+)",
    re.IGNORECASE,
)
_LOCAL_CELL_ABS_RE = re.compile(
    r"(?<![!A-Za-z0-9_])(?<!\$)(?P<acol>\$)?(?P<col>[A-Z]{1,3})(?P<arow>\$)?(?P<row>\d+)"
    r"(?![A-Za-z0-9_!'])",
    re.IGNORECASE,
)
_RANGE_LOCAL_RE = re.compile(
    r"(?<![!A-Za-z0-9_])(?<!\$)"
    r"(?P<c1a>\$)?(?P<c1>[A-Z]{1,3})(?P<r1a>\$)?(?P<r1>\d+)\s*:\s*"
    r"(?P<c2a>\$)?(?P<c2>[A-Z]{1,3})(?P<r2a>\$)?(?P<r2>\d+)"
    r"(?![A-Za-z0-9_])",
    re.IGNORECASE,
)
_FUNC_LIKE = {"IF", "OR", "AND", "NOT", "SUM", "MAX", "MIN", "AVG"}


def parse_ref_streams(formula: str, *, default_sheet: str) -> list[FormulaRef]:
    """Extract formula references in source order (cells and ranges)."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    streams: list[FormulaRef] = []
    range_spans: list[tuple[int, int]] = []

    for _start, _end, span in parse_range_refs_with_spans(formula):
        text = formula[span[0] : span[1]]
        parsed = _parse_range_text(text, default_sheet=default_sheet)
        if parsed is not None:
            sheet, sc, sr, ec, er, s_ac, s_ar, e_ac, e_ar = parsed
            streams.append(
                AbsRangeRef(
                    sheet=sheet,
                    start_col=sc,
                    start_row=sr,
                    end_col=ec,
                    end_row=er,
                    start_abs_col=s_ac,
                    start_abs_row=s_ar,
                    end_abs_col=e_ac,
                    end_abs_row=e_ar,
                    span=span,
                )
            )
        range_spans.append(span)

    masked = mask_spans(formula, range_spans)
    for m in _SHEET_CELL_ABS_RE.finditer(masked):
        sheet = m.group("qs") or m.group("us")
        streams.append(
            AbsCellRef(
                sheet=sheet,
                column=m.group("col").upper(),
                row=int(m.group("row")),
                is_absolute_col=m.group("acol") is not None,
                is_absolute_row=m.group("arow") is not None,
                span=m.span(),
            )
        )

    for m in _LOCAL_CELL_ABS_RE.finditer(masked):
        col = m.group("col").upper()
        if col in _FUNC_LIKE:
            continue
        streams.append(
            AbsCellRef(
                sheet=None,
                column=col,
                row=int(m.group("row")),
                is_absolute_col=m.group("acol") is not None,
                is_absolute_row=m.group("arow") is not None,
                span=m.span(),
            )
        )

    streams.sort(key=lambda ref: ref.span[0])
    return streams


def parse_cell_refs_with_abs(formula: str, *, default_sheet: str) -> list[AbsCellRef]:
    """Extract single-cell references from a raw formula, preserving `$` markers."""
    return [
        ref
        for ref in parse_ref_streams(formula, default_sheet=default_sheet)
        if isinstance(ref, AbsCellRef)
    ]


def abs_ref_to_key(ref: AbsCellRef, *, default_sheet: str) -> str:
    """Resolve an absolute cell ref to a sheet-qualified graph key."""
    from excel_grapher.core.address_keys import format_cell_key

    sheet = ref.sheet if ref.sheet is not None else default_sheet
    return format_cell_key(sheet, ref.column, ref.row)


def range_ref_to_keys(ref: AbsRangeRef, *, default_sheet: str) -> list[str]:
    """Expand a parsed range reference to sheet-qualified cell keys."""
    sheet = ref.sheet if ref.sheet is not None else default_sheet
    return expand_resolved_range(sheet, ref.start_col, ref.start_row, ref.end_col, ref.end_row)


def expand_resolved_range(
    sheet: str,
    start_col: str,
    start_row: int,
    end_col: str,
    end_row: int,
) -> list[str]:
    """Expand a resolved rectangular range to sheet-qualified keys."""
    import fastpyxl.utils.cell as xl_cell

    from excel_grapher.core.address_keys import format_cell_key

    c1 = xl_cell.column_index_from_string(start_col)
    c2 = xl_cell.column_index_from_string(end_col)
    r1, r2 = (start_row, end_row) if start_row <= end_row else (end_row, start_row)
    clo, chi = (c1, c2) if c1 <= c2 else (c2, c1)
    out: list[str] = []
    for row in range(r1, r2 + 1):
        for col_i in range(clo, chi + 1):
            out.append(format_cell_key(sheet, xl_cell.get_column_letter(col_i), row))
    return out


def _parse_range_text(
    text: str, *, default_sheet: str
) -> tuple[str | None, str, int, str, int, bool, bool, bool, bool] | None:
    m = re.match(
        r"^(?:'(?P<qs>[^']+)'|(?P<us>[A-Za-z][A-Za-z0-9_]*))!"
        r"(?P<c1a>\$)?(?P<c1>[A-Z]{1,3})(?P<r1a>\$)?(?P<r1>\d+)\s*:\s*"
        r"(?:(?:'(?P=qs)'|(?P=us))!)?"
        r"(?P<c2a>\$)?(?P<c2>[A-Z]{1,3})(?P<r2a>\$)?(?P<r2>\d+)$",
        text,
        re.IGNORECASE,
    )
    if m:
        sheet = m.group("qs") or m.group("us")
        return (
            sheet,
            m.group("c1").upper(),
            int(m.group("r1")),
            m.group("c2").upper(),
            int(m.group("r2")),
            m.group("c1a") is not None,
            m.group("r1a") is not None,
            m.group("c2a") is not None,
            m.group("r2a") is not None,
        )
    m = _RANGE_LOCAL_RE.match(text)
    if m:
        return (
            None,
            m.group("c1").upper(),
            int(m.group("r1")),
            m.group("c2").upper(),
            int(m.group("r2")),
            m.group("c1a") is not None,
            m.group("r1a") is not None,
            m.group("c2a") is not None,
            m.group("r2a") is not None,
        )
    return None


def classify_range_pattern(ref: AbsRangeRef) -> str | None:
    """Return RF, FR, or FF when the endpoint `$` pattern matches TACO classes."""
    start_rel = not ref.start_abs_col and not ref.start_abs_row
    start_fixed = ref.start_abs_col and ref.start_abs_row
    end_rel = not ref.end_abs_col and not ref.end_abs_row
    end_fixed = ref.end_abs_col and ref.end_abs_row
    if start_rel and end_fixed:
        return "RF"
    if start_fixed and end_rel:
        return "FR"
    if start_fixed and end_fixed:
        return "FF"
    if start_rel and end_rel:
        return "RR"
    return None
