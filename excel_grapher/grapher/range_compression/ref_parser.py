"""Parse cell references from raw formulas, preserving `$` absolute markers."""

from __future__ import annotations

import re
from dataclasses import dataclass

from excel_grapher.grapher.parser import mask_spans, parse_range_refs_with_spans


@dataclass(frozen=True, slots=True)
class AbsCellRef:
    """Single cell reference extracted from a raw formula."""

    sheet: str | None
    column: str
    row: int
    is_absolute_col: bool
    is_absolute_row: bool


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
_FUNC_LIKE = {"IF", "OR", "AND", "NOT", "SUM", "MAX", "MIN", "AVG"}


def parse_cell_refs_with_abs(formula: str, *, default_sheet: str) -> list[AbsCellRef]:
    """Extract single-cell references from a raw formula, preserving `$` markers.

    Range endpoints are masked so they are not double-counted as separate cell refs.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    range_spans = [span for _s, _e, span in parse_range_refs_with_spans(formula)]
    masked = mask_spans(formula, range_spans)

    out: list[AbsCellRef] = []
    for m in _SHEET_CELL_ABS_RE.finditer(masked):
        sheet = m.group("qs") or m.group("us")
        out.append(
            AbsCellRef(
                sheet=sheet,
                column=m.group("col").upper(),
                row=int(m.group("row")),
                is_absolute_col=m.group("acol") is not None,
                is_absolute_row=m.group("arow") is not None,
            )
        )

    for m in _LOCAL_CELL_ABS_RE.finditer(masked):
        col = m.group("col").upper()
        if col in _FUNC_LIKE:
            continue
        out.append(
            AbsCellRef(
                sheet=None,
                column=col,
                row=int(m.group("row")),
                is_absolute_col=m.group("acol") is not None,
                is_absolute_row=m.group("arow") is not None,
            )
        )

    return out


def abs_ref_to_key(ref: AbsCellRef, *, default_sheet: str) -> str:
    """Resolve an absolute cell ref to a sheet-qualified graph key."""
    from excel_grapher.core.address_keys import format_cell_key

    sheet = ref.sheet if ref.sheet is not None else default_sheet
    return format_cell_key(sheet, ref.column, ref.row)
