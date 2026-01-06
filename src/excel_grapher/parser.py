from __future__ import annotations

import re
from dataclasses import dataclass

import openpyxl.utils.cell


@dataclass(frozen=True)
class CellRef:
    sheet: str | None
    column: str
    row: int
    is_absolute_col: bool = False
    is_absolute_row: bool = False


_SHEET_CELL_RE = re.compile(
    r"(?:'(?P<qs>[^']+)'|(?P<us>[A-Za-z][A-Za-z0-9_]*))!\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)"
)
_LOCAL_CELL_RE = re.compile(
    r"(?<![!A-Za-z0-9_])\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)(?![A-Za-z0-9_])"
)
_FUNC_LIKE = {"IF", "OR", "AND", "NOT", "SUM", "MAX", "MIN", "AVG"}

_RANGE_QUOTED_RE = re.compile(
    r"'(?P<sheet>[^']+)'!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)"
)
_RANGE_UNQUOTED_RE = re.compile(
    r"(?<![A-Za-z_])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)"
)
_RANGE_LOCAL_RE = re.compile(
    r"(?<![!A-Za-z0-9_])\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)(?![A-Za-z0-9_])"
)


def parse_cell_refs(formula: str) -> list[CellRef]:
    """
    Extract single-cell references from a formula.

    This function does not expand ranges. Use parse_range_refs + expand_range.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    out: list[CellRef] = []

    for m in _SHEET_CELL_RE.finditer(formula):
        sheet = m.group("qs") or m.group("us")
        out.append(CellRef(sheet=sheet, column=m.group("col"), row=int(m.group("row"))))

    for m in _LOCAL_CELL_RE.finditer(formula):
        col = m.group("col")
        if col in _FUNC_LIKE:
            continue
        out.append(CellRef(sheet=None, column=col, row=int(m.group("row"))))

    return out


def parse_range_refs(formula: str) -> list[tuple[CellRef, CellRef]]:
    """
    Extract range references from a formula as (start, end) CellRef pairs.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    out: list[tuple[CellRef, CellRef]] = []

    for m in _RANGE_QUOTED_RE.finditer(formula):
        sheet = m.group("sheet")
        out.append(
            (
                CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
            )
        )

    for m in _RANGE_UNQUOTED_RE.finditer(formula):
        sheet = m.group("sheet")
        out.append(
            (
                CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
            )
        )

    for m in _RANGE_LOCAL_RE.finditer(formula):
        out.append(
            (
                CellRef(sheet=None, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=None, column=m.group("c2"), row=int(m.group("r2"))),
            )
        )

    return out


def parse_range_refs_with_spans(formula: str) -> list[tuple[CellRef, CellRef, tuple[int, int]]]:
    """
    Extract range references from a formula as (start, end, (span_start, span_end)).

    The span corresponds to the exact matched substring in the original formula and is
    intended for masking to prevent endpoint tokens from being re-parsed as separate refs.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    out: list[tuple[CellRef, CellRef, tuple[int, int]]] = []

    for m in _RANGE_QUOTED_RE.finditer(formula):
        sheet = m.group("sheet")
        out.append(
            (
                CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
                m.span(),
            )
        )

    for m in _RANGE_UNQUOTED_RE.finditer(formula):
        sheet = m.group("sheet")
        out.append(
            (
                CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
                m.span(),
            )
        )

    for m in _RANGE_LOCAL_RE.finditer(formula):
        out.append(
            (
                CellRef(sheet=None, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=None, column=m.group("c2"), row=int(m.group("r2"))),
                m.span(),
            )
        )

    return out


def expand_range(
    *,
    sheet: str,
    start_col: str,
    start_row: int,
    end_col: str,
    end_row: int,
    max_cells: int,
) -> list[tuple[str, str]]:
    """
    Expand an A1 range into individual (sheet, A1) dependencies.
    """
    c1i = openpyxl.utils.cell.column_index_from_string(start_col)
    c2i = openpyxl.utils.cell.column_index_from_string(end_col)
    rlo, rhi = (start_row, end_row) if start_row <= end_row else (end_row, start_row)
    clo, chi = (c1i, c2i) if c1i <= c2i else (c2i, c1i)
    n_cells = (rhi - rlo + 1) * (chi - clo + 1)
    if n_cells > max_cells:
        return [(sheet, f"{start_col}{start_row}"), (sheet, f"{end_col}{end_row}")]

    out: list[tuple[str, str]] = []
    for rr in range(rlo, rhi + 1):
        for cc in range(clo, chi + 1):
            out.append((sheet, f"{openpyxl.utils.cell.get_column_letter(cc)}{rr}"))
    return out


def mask_spans(s: str, spans: list[tuple[int, int]]) -> str:
    """
    Replace characters in the specified spans with spaces.

    Useful to prevent endpoint tokens from being re-parsed as separate refs.
    """
    if not spans:
        return s
    buf = list(s)
    for a, b in spans:
        for i in range(a, b):
            buf[i] = " "
    return "".join(buf)


def _needs_quoting(sheet: str) -> bool:
    """Return True if sheet name needs quoting in a formula."""
    # Sheets with spaces or special chars need quoting
    return " " in sheet or "-" in sheet or "'" in sheet


def _format_ref(sheet: str, col: str, row: int) -> str:
    """Format a fully-qualified cell reference."""
    if _needs_quoting(sheet):
        return f"'{sheet}'!{col}{row}"
    return f"{sheet}!{col}{row}"


def normalize_formula(
    formula: str,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]] | None = None,
) -> str:
    """
    Normalize a formula for transpilation:
    
    - Replace same-sheet refs (A1) with sheet-qualified refs (Sheet1!A1)
    - Resolve named ranges to their targets
    - Strip absolute markers ($)
    - Qualify range endpoints
    
    Returns the normalized formula string.
    """
    if not formula or not formula.startswith("="):
        return formula
    
    if named_ranges is None:
        named_ranges = {}
    
    result = formula
    
    # 1) Normalize ranges first (quoted sheet, unquoted sheet, local)
    # Pattern for 'Sheet Name'!$A$1:$B$2
    def replace_quoted_range(m: re.Match) -> str:
        sheet = m.group("sheet")
        c1 = m.group("c1")
        r1 = m.group("r1")
        c2 = m.group("c2")
        r2 = m.group("r2")
        return f"'{sheet}'!{c1}{r1}:'{sheet}'!{c2}{r2}"
    
    result = re.sub(
        r"'(?P<sheet>[^']+)'!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_quoted_range,
        result,
    )
    
    # Pattern for SheetName!$A$1:$B$2 (unquoted)
    def replace_unquoted_range(m: re.Match) -> str:
        sheet = m.group("sheet")
        c1 = m.group("c1")
        r1 = m.group("r1")
        c2 = m.group("c2")
        r2 = m.group("r2")
        return f"{sheet}!{c1}{r1}:{sheet}!{c2}{r2}"
    
    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_unquoted_range,
        result,
    )
    
    # Pattern for local range A1:B2 -> Sheet!A1:Sheet!B2
    def replace_local_range(m: re.Match) -> str:
        c1 = m.group("c1")
        r1 = m.group("r1")
        c2 = m.group("c2")
        r2 = m.group("r2")
        ref1 = _format_ref(current_sheet, c1, int(r1))
        ref2 = _format_ref(current_sheet, c2, int(r2))
        return f"{ref1}:{ref2}"
    
    result = re.sub(
        r"(?<![!A-Za-z0-9_])\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)(?![A-Za-z0-9_])",
        replace_local_range,
        result,
    )
    
    # 2) Normalize sheet-qualified single-cell refs (strip $)
    # 'Sheet Name'!$A$1 -> 'Sheet Name'!A1
    def replace_quoted_cell(m: re.Match) -> str:
        sheet = m.group("sheet")
        col = m.group("col")
        row = m.group("row")
        return f"'{sheet}'!{col}{row}"
    
    result = re.sub(
        r"'(?P<sheet>[^']+)'!\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)",
        replace_quoted_cell,
        result,
    )
    
    # SheetName!$A$1 -> SheetName!A1
    def replace_unquoted_cell(m: re.Match) -> str:
        sheet = m.group("sheet")
        col = m.group("col")
        row = m.group("row")
        return f"{sheet}!{col}{row}"
    
    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)",
        replace_unquoted_cell,
        result,
    )
    
    # 3) Normalize local single-cell refs: $A$1, A$1, $A1, A1 -> Sheet!A1
    def replace_local_cell(m: re.Match) -> str:
        col = m.group("col")
        row = m.group("row")
        # Skip function-like tokens
        if col in _FUNC_LIKE:
            return m.group(0)
        return _format_ref(current_sheet, col, int(row))
    
    result = re.sub(
        r"(?<![!A-Za-z0-9_])\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)(?![A-Za-z0-9_])",
        replace_local_cell,
        result,
    )
    
    # 4) Resolve named ranges
    def replace_named_range(m: re.Match) -> str:
        token = m.group(1)
        if token in named_ranges:
            sheet, addr = named_ranges[token]
            return _format_ref(sheet, addr[:-len(addr.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))] or addr, int(addr.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ")) if addr.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ").isdigit() else 0)
        return token
    
    # Actually, let's do a simpler approach for named ranges
    for name, (sheet, addr) in named_ranges.items():
        # Parse the address to get col and row
        col_match = re.match(r"([A-Z]+)(\d+)", addr)
        if col_match:
            col = col_match.group(1)
            row = int(col_match.group(2))
            replacement = _format_ref(sheet, col, row)
            # Replace whole-word occurrences of the name
            result = re.sub(rf"\b{re.escape(name)}\b(?!\s*!)", replacement, result)
    
    return result

