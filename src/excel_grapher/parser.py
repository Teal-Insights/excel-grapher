from __future__ import annotations

import re
from dataclasses import dataclass

import openpyxl.utils.cell

from .guard import And, Compare, GuardExpr, Literal, Not, Or
from .guard import CellRef as GuardCellRef
from .node import NodeKey


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


def needs_quoting(sheet: str) -> bool:
    """Return True if sheet name needs quoting in a formula."""
    # Sheets with spaces or special chars need quoting
    return " " in sheet or "-" in sheet or "'" in sheet


def format_cell_key(sheet: str, col: str, row: int) -> str:
    """
    Format a fully-qualified cell reference key.

    Quotes the sheet name if it contains spaces, hyphens, or apostrophes.
    This matches Excel's formula syntax for sheet references.

    Examples:
        >>> format_cell_key("Sheet1", "A", 1)
        'Sheet1!A1'
        >>> format_cell_key("My Sheet", "B", 2)
        "'My Sheet'!B2"
        >>> format_cell_key("Baseline - external", "M", 35)
        "'Baseline - external'!M35"
    """
    if needs_quoting(sheet):
        return f"'{sheet}'!{col}{row}"
    return f"{sheet}!{col}{row}"


def format_key(sheet: str, a1: str) -> str:
    """
    Format a fully-qualified cell key from sheet name and A1 address.

    Quotes the sheet name if it contains spaces, hyphens, or apostrophes.
    This matches Excel's formula syntax for sheet references.

    Examples:
        >>> format_key("Sheet1", "A1")
        'Sheet1!A1'
        >>> format_key("My Sheet", "B2")
        "'My Sheet'!B2"
    """
    if needs_quoting(sheet):
        return f"'{sheet}'!{a1}"
    return f"{sheet}!{a1}"


# Alias for internal use
_needs_quoting = needs_quoting
_format_ref = format_cell_key


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


def split_top_level_if(formula: str) -> tuple[str, str, str] | None:
    """
    If formula is a top-level IF(...), return (cond, then_expr, else_expr) strings.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    s = formula[1:].lstrip()
    if s[:3].upper() != "IF(":
        return None

    # Parse the IF argument list at the top level of IF(...).
    i = s.find("(")
    if i < 0:
        return None
    inner = s[i + 1 :]

    args: list[str] = []
    buf: list[str] = []
    depth = 0
    in_str = False
    j = 0
    while j < len(inner):
        ch = inner[j]
        if ch == '"':
            in_str = not in_str
            buf.append(ch)
            j += 1
            continue
        if in_str:
            buf.append(ch)
            j += 1
            continue
        if ch == "(":
            depth += 1
            buf.append(ch)
            j += 1
            continue
        if ch == ")":
            if depth == 0:
                args.append("".join(buf).strip())
                buf = []
                # Consume the closing paren and stop; ignore any trailing whitespace.
                j += 1
                break
            depth -= 1
            buf.append(ch)
            j += 1
            continue
        if ch == "," and depth == 0:
            args.append("".join(buf).strip())
            buf = []
            j += 1
            continue
        buf.append(ch)
        j += 1

    # Check for trailing content - if present, this is not a top-level IF
    remaining = inner[j:].strip()
    if remaining:
        return None

    if len(args) != 3:
        return None
    cond, then_expr, else_expr = args
    if not cond or not then_expr:
        return None
    # Excel allows empty else, treat as "" expression (no deps).
    return cond, then_expr, else_expr


def split_top_level_function(formula: str, fn: str) -> list[str] | None:
    """
    If formula is a top-level FN(...), return the top-level argument strings.

    This is a lightweight splitter that is aware of nested parentheses and quoted
    string literals (") so commas inside those don't split arguments.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    s = formula[1:].lstrip()
    fn_u = fn.upper()
    prefixes = (f"{fn_u}(", f"_XLFN.{fn_u}(")
    if not any(s[: len(p)].upper() == p for p in prefixes):
        return None

    # Parse the argument list at the top level of FN(...).
    i = s.find("(")
    if i < 0:
        return None
    inner = s[i + 1 :]

    args: list[str] = []
    buf: list[str] = []
    depth = 0
    in_str = False
    j = 0
    while j < len(inner):
        ch = inner[j]
        if ch == '"':
            in_str = not in_str
            buf.append(ch)
            j += 1
            continue
        if in_str:
            buf.append(ch)
            j += 1
            continue
        if ch == "(":
            depth += 1
            buf.append(ch)
            j += 1
            continue
        if ch == ")":
            if depth == 0:
                args.append("".join(buf).strip())
                buf = []
                j += 1
                break
            depth -= 1
            buf.append(ch)
            j += 1
            continue
        if ch == "," and depth == 0:
            args.append("".join(buf).strip())
            buf = []
            j += 1
            continue
        buf.append(ch)
        j += 1

    # Check for trailing content - if present, this is not a top-level function
    remaining = inner[j:].strip()
    if remaining:
        return None

    if in_str or depth != 0:
        return None
    return args


def split_top_level_ifs(formula: str) -> list[str] | None:
    """
    If formula is a top-level IFS(...), return argument strings.
    """
    return split_top_level_function(formula, "IFS")


def split_top_level_choose(formula: str) -> list[str] | None:
    """
    If formula is a top-level CHOOSE(...), return argument strings.
    """
    return split_top_level_function(formula, "CHOOSE")


def split_top_level_switch(formula: str) -> list[str] | None:
    """
    If formula is a top-level SWITCH(...), return argument strings.
    """
    return split_top_level_function(formula, "SWITCH")


_CELL_TOKEN_RE = re.compile(
    r"^(?:(?:'(?P<qs>[^']+)')|(?P<us>[A-Za-z][A-Za-z0-9_]*))!\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)$"
)
_LOCAL_CELL_TOKEN_RE = re.compile(r"^\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)$")
_NUMBER_TOKEN_RE = re.compile(r"^-?\d+(?:\.\d+)?$")


def _to_node_key(sheet: str, col: str, row: int) -> NodeKey:
    return format_cell_key(sheet, col, row)


def parse_guard_expr(
    expr: str,
    *,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]] | None = None,
) -> GuardExpr | None:
    """
    Parse a minimal boolean expression suitable for IF(...) conditions.

    If parsing fails (unsupported syntax), returns None.
    """
    if named_ranges is None:
        named_ranges = {}
    s = expr.strip()
    if not s:
        return None

    # Strip redundant outer parentheses.
    while s.startswith("(") and s.endswith(")"):
        inner = s[1:-1].strip()
        if not inner:
            break
        s = inner

    # Function-like: AND(...), OR(...), NOT(...)
    m = re.match(r"^(?P<fn>AND|OR|NOT)\s*\((?P<inner>.*)\)$", s, flags=re.IGNORECASE)
    if m:
        fn = m.group("fn").upper()
        inner = m.group("inner")
        parts = _split_top_level_args(inner)
        if parts is None:
            return None
        if fn == "NOT":
            if len(parts) != 1:
                return None
            operand = parse_guard_expr(parts[0], current_sheet=current_sheet, named_ranges=named_ranges)
            return None if operand is None else Not(operand)
        if fn in {"AND", "OR"}:
            if len(parts) < 1:
                return None
            ops: list[GuardExpr] = []
            for p in parts:
                ge = parse_guard_expr(p, current_sheet=current_sheet, named_ranges=named_ranges)
                if ge is None:
                    return None
                ops.append(ge)
            return And(tuple(ops)) if fn == "AND" else Or(tuple(ops))

    # Comparison: left op right
    for op in ("<>", "<=", ">=", "=", "<", ">"):
        if op in s:
            left_s, right_s = (p.strip() for p in s.split(op, 1))
            left = _parse_guard_atom(left_s, current_sheet=current_sheet, named_ranges=named_ranges)
            right = _parse_guard_atom(right_s, current_sheet=current_sheet, named_ranges=named_ranges)
            if left is None or right is None:
                return None
            return Compare(left=left, op=op, right=right)

    # Atom (truthy cell ref or literal)
    return _parse_guard_atom(s, current_sheet=current_sheet, named_ranges=named_ranges)


def _split_top_level_args(s: str) -> list[str] | None:
    buf: list[str] = []
    args: list[str] = []
    depth = 0
    in_str = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '"':
            in_str = not in_str
            buf.append(ch)
            i += 1
            continue
        if in_str:
            buf.append(ch)
            i += 1
            continue
        if ch == "(":
            depth += 1
            buf.append(ch)
            i += 1
            continue
        if ch == ")":
            if depth == 0:
                return None
            depth -= 1
            buf.append(ch)
            i += 1
            continue
        if ch == "," and depth == 0:
            args.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    if in_str or depth != 0:
        return None
    args.append("".join(buf).strip())
    return [a for a in args if a != ""]


def _parse_guard_atom(
    s: str,
    *,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]],
) -> GuardExpr | None:
    if not s:
        return None

    # TRUE/FALSE
    if s.upper() == "TRUE":
        return Literal(True)
    if s.upper() == "FALSE":
        return Literal(False)

    # Number literal
    if _NUMBER_TOKEN_RE.match(s):
        if "." in s:
            try:
                return Literal(float(s))
            except ValueError:
                return None
        try:
            return Literal(int(s))
        except ValueError:
            return None

    # String literal
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        return Literal(s[1:-1])

    # Named range (single-cell only)
    if s in named_ranges:
        sheet, addr = named_ranges[s]
        m2 = re.match(r"^(?P<col>[A-Z]{1,3})(?P<row>\d+)$", addr)
        if not m2:
            return None
        return GuardCellRef(_to_node_key(sheet, m2.group("col"), int(m2.group("row"))))

    # Sheet-qualified cell
    m = _CELL_TOKEN_RE.match(s)
    if m:
        sheet = m.group("qs") or m.group("us")
        return GuardCellRef(_to_node_key(str(sheet), m.group("col"), int(m.group("row"))))

    # Local cell
    m = _LOCAL_CELL_TOKEN_RE.match(s)
    if m:
        col = m.group("col")
        if col in _FUNC_LIKE:
            return None
        return GuardCellRef(_to_node_key(current_sheet, col, int(m.group("row"))))

    return None


