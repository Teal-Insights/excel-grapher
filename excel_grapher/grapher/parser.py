from __future__ import annotations

import functools
import re
import warnings
from collections.abc import Callable
from dataclasses import dataclass
from typing import Literal as TypingLiteral

import fastpyxl.utils.cell

from excel_grapher.core import address_keys as _address_keys
from excel_grapher.core.address_keys import (
    format_cell_key,
    format_range_key,
    needs_quoting,
    quoted_sheet_name_regex,
    quoted_sheet_prefix_regex,
    unescape_formula_sheet_name,
)
from excel_grapher.core.excel_function_meta import ref_only_function_names
from excel_grapher.core.excel_function_names import excel_function_call_prefixes
from excel_grapher.core.formula_normalization import (
    build_named_range_replacement_state,
    normalize_excel_formula,
    normalize_excel_formula_with_name_state,
)
from excel_grapher.core.range_shorthand import (
    SheetBounds,
    expand_whole_column_deps,
    expand_whole_row_deps,
)

from .guard import And, Compare, GuardExpr, Literal, Not, Or, RangeRef, intern_guard
from .guard import CellRef as GuardCellRef
from .node import NodeKey

# Re-exported for `from excel_grapher.grapher.parser import format_key, ...`
format_key = _address_keys.format_key
normalize_key = _address_keys.normalize_key
parse_address = _address_keys.parse_address


@dataclass(frozen=True)
class CellRef:
    sheet: str | None
    column: str
    row: int
    is_absolute_col: bool = False
    is_absolute_row: bool = False
    range_kind: TypingLiteral["cell", "whole_column", "whole_row"] = "cell"


_FUNC_LIKE = {"IF", "OR", "AND", "NOT", "SUM", "MAX", "MIN", "AVG"}

_QUOTED_SHEET = quoted_sheet_name_regex(capture_group="sheet")
_QUOTED_SHEET_QS = quoted_sheet_name_regex(capture_group="qs")
_QUOTED_SHEET_PREFIX = quoted_sheet_prefix_regex()


def _sheet_from_cell_match(m: re.Match[str]) -> str:
    qs = m.group("qs")
    if qs is not None:
        return unescape_formula_sheet_name(qs)
    us = m.group("us")
    assert us is not None
    return us


def _sheet_from_quoted_group(m: re.Match[str], *, group: str = "sheet") -> str:
    return unescape_formula_sheet_name(m.group(group))


_SHEET_CELL_RE = re.compile(
    rf"(?:{_QUOTED_SHEET_QS}|(?P<us>[A-Za-z][A-Za-z0-9_]*))!\$?(?P<col>[A-Z]{{1,3}})\$?(?P<row>\d+)"
)
# Do not treat `Col` in `Sheet!$Col$Row` as a local ref: the column letter can follow `!$`.
# Sheet tokens may look like A1 (e.g. `S2` / `'S2'`); do not match those as local refs.
# Exclude `:` so bare endpoints in single-prefix ranges (`Sheet1!A1:A3`) are not
# treated as local cells; callers that also expand ranges should still mask spans.
_LOCAL_CELL_RE = re.compile(
    r"(?<![!A-Za-z0-9_:])(?<!\$)\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)(?![A-Za-z0-9_!'])"
)

_RANGE_QUOTED_RE = re.compile(
    _QUOTED_SHEET_PREFIX
    + r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)"
)
_RANGE_UNQUOTED_RE = re.compile(
    r"(?<![A-Za-z_])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)"
)
_RANGE_QUOTED_BOTH_ENDPOINTS_RE = re.compile(
    _QUOTED_SHEET_PREFIX
    + r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*'(?P=sheet)'!\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)"
)
_RANGE_UNQUOTED_BOTH_ENDPOINTS_RE = re.compile(
    r"(?<![A-Za-z_])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*(?P=sheet)!\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)"
)
_RANGE_LOCAL_RE = re.compile(
    r"(?<![!A-Za-z0-9_])(?<!\$)\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)(?![A-Za-z0-9_])"
)
_RANGE_WHOLE_COL_QUOTED_RE = re.compile(
    _QUOTED_SHEET_PREFIX + r"\$?(?P<col>[A-Z]{1,3})\s*:\s*\$?(?P=col)\b",
    re.IGNORECASE,
)
_RANGE_WHOLE_COL_UNQUOTED_RE = re.compile(
    r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<col>[A-Z]{1,3})\s*:\s*\$?(?P=col)\b",
    re.IGNORECASE,
)
_RANGE_WHOLE_COL_LOCAL_RE = re.compile(
    r"(?<![!A-Za-z0-9_'])(?<!\$)\$?(?P<col>[A-Z]{1,3})\s*:\s*\$?(?P=col)\b(?![A-Za-z0-9_])",
    re.IGNORECASE,
)
_RANGE_WHOLE_ROW_QUOTED_RE = re.compile(
    _QUOTED_SHEET_PREFIX + r"\$?(?P<row>\d+)\s*:\s*\$?(?P=row)\b"
)
_RANGE_WHOLE_ROW_UNQUOTED_RE = re.compile(
    r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<row>\d+)\s*:\s*\$?(?P=row)\b"
)
_RANGE_WHOLE_ROW_LOCAL_RE = re.compile(
    r"(?<![!A-Za-z0-9_'])(?<!\$)(?P<row>\d+)\s*:\s*\$?(?P=row)\b(?![A-Za-z0-9_])"
)


def _reject_unmasked_ranges(formula: str, *, allow_unmasked_ranges: bool) -> None:
    """Raise if *formula* still contains range spans and opt-in was not given."""
    if allow_unmasked_ranges:
        return
    if parse_range_refs_with_spans(formula):
        raise ValueError(
            "Formula text still contains unmasked range spans; "
            "prefer parse_standalone_cell_refs (or mask_spans) so range endpoints "
            "are not treated as standalone cells, or pass allow_unmasked_ranges=True"
        )


def parse_cell_refs(formula: str, *, allow_unmasked_ranges: bool = False) -> list[CellRef]:
    """Extract single-cell references from a formula.

    This function does not expand ranges. Use parse_range_refs + expand_range.
    By default it refuses formula text that still contains range spans — prefer
    `parse_standalone_cell_refs` (or mask spans yourself). Pass
    `allow_unmasked_ranges=True` only when intentionally inspecting raw text
    that may include ranges (for example lookbehind regression tests).
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    _reject_unmasked_ranges(formula, allow_unmasked_ranges=allow_unmasked_ranges)

    out: list[CellRef] = []

    for m in _SHEET_CELL_RE.finditer(formula):
        sheet = _sheet_from_cell_match(m)
        out.append(CellRef(sheet=sheet, column=m.group("col"), row=int(m.group("row"))))

    for m in _LOCAL_CELL_RE.finditer(formula):
        col = m.group("col")
        if col in _FUNC_LIKE:
            continue
        out.append(CellRef(sheet=None, column=col, row=int(m.group("row"))))

    return out


def parse_standalone_cell_refs(formula: str) -> list[CellRef]:
    """Extract cell refs after masking range spans.

    Prefer this over `parse_cell_refs` whenever ranges may appear in the same
    text, so single-prefix forms like `Sheet1!A1:A3` do not yield a bare `A3`
    or double-count the range start as a sheet-qualified cell.
    """
    spans = [span for _start, _end, span in parse_range_refs_with_spans(formula)]
    return parse_cell_refs(mask_spans(formula, spans))


def parse_cell_refs_with_spans(
    formula: str, *, allow_unmasked_ranges: bool = False
) -> list[tuple[CellRef, tuple[int, int]]]:
    """Like parse_cell_refs, but returns (CellRef, (start, end)) span positions in `formula`."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    _reject_unmasked_ranges(formula, allow_unmasked_ranges=allow_unmasked_ranges)

    out: list[tuple[CellRef, tuple[int, int]]] = []

    for m in _SHEET_CELL_RE.finditer(formula):
        sheet = _sheet_from_cell_match(m)
        ref = CellRef(sheet=sheet, column=m.group("col"), row=int(m.group("row")))
        out.append((ref, m.span()))

    for m in _LOCAL_CELL_RE.finditer(formula):
        col = m.group("col")
        if col in _FUNC_LIKE:
            continue
        ref = CellRef(sheet=None, column=col, row=int(m.group("row")))
        out.append((ref, m.span()))

    return out


def parse_standalone_cell_refs_with_spans(
    formula: str,
) -> list[tuple[CellRef, tuple[int, int]]]:
    """Like `parse_standalone_cell_refs`, but keep character spans in `formula`.

    Spans remain valid in the original text because `mask_spans` only replaces
    range characters with spaces of the same width.
    """
    spans = [span for _start, _end, span in parse_range_refs_with_spans(formula)]
    return parse_cell_refs_with_spans(mask_spans(formula, spans))


def parse_range_refs(formula: str) -> list[tuple[CellRef, CellRef]]:
    """Extract range references from a formula as (start, end) CellRef pairs."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    out: list[tuple[CellRef, CellRef]] = []

    for m in _RANGE_WHOLE_COL_QUOTED_RE.finditer(formula):
        sheet = _sheet_from_quoted_group(m)
        col = m.group("col").upper()
        ref = CellRef(sheet=sheet, column=col, row=0, range_kind="whole_column")
        out.append((ref, ref))

    for m in _RANGE_WHOLE_COL_UNQUOTED_RE.finditer(formula):
        sheet = m.group("sheet")
        col = m.group("col").upper()
        ref = CellRef(sheet=sheet, column=col, row=0, range_kind="whole_column")
        out.append((ref, ref))

    for m in _RANGE_WHOLE_COL_LOCAL_RE.finditer(formula):
        col = m.group("col").upper()
        ref = CellRef(sheet=None, column=col, row=0, range_kind="whole_column")
        out.append((ref, ref))

    for m in _RANGE_WHOLE_ROW_QUOTED_RE.finditer(formula):
        sheet = _sheet_from_quoted_group(m)
        row = int(m.group("row"))
        ref = CellRef(sheet=sheet, column="", row=row, range_kind="whole_row")
        out.append((ref, ref))

    for m in _RANGE_WHOLE_ROW_UNQUOTED_RE.finditer(formula):
        sheet = m.group("sheet")
        row = int(m.group("row"))
        ref = CellRef(sheet=sheet, column="", row=row, range_kind="whole_row")
        out.append((ref, ref))

    for m in _RANGE_WHOLE_ROW_LOCAL_RE.finditer(formula):
        row = int(m.group("row"))
        ref = CellRef(sheet=None, column="", row=row, range_kind="whole_row")
        out.append((ref, ref))

    for m in _RANGE_QUOTED_BOTH_ENDPOINTS_RE.finditer(formula):
        sheet = _sheet_from_quoted_group(m)
        out.append(
            (
                CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
            )
        )

    for m in _RANGE_UNQUOTED_BOTH_ENDPOINTS_RE.finditer(formula):
        sheet = m.group("sheet")
        out.append(
            (
                CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
            )
        )

    for m in _RANGE_QUOTED_RE.finditer(formula):
        sheet = _sheet_from_quoted_group(m)
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
    """Extract range references from a formula as (start, end, (span_start, span_end)).

    The span corresponds to the exact matched substring in the original formula and is
    intended for masking to prevent endpoint tokens from being re-parsed as separate refs.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    out: list[tuple[CellRef, CellRef, tuple[int, int]]] = []

    for m in _RANGE_WHOLE_COL_QUOTED_RE.finditer(formula):
        sheet = _sheet_from_quoted_group(m)
        col = m.group("col").upper()
        ref = CellRef(sheet=sheet, column=col, row=0, range_kind="whole_column")
        out.append((ref, ref, m.span()))

    for m in _RANGE_WHOLE_COL_UNQUOTED_RE.finditer(formula):
        sheet = m.group("sheet")
        col = m.group("col").upper()
        ref = CellRef(sheet=sheet, column=col, row=0, range_kind="whole_column")
        out.append((ref, ref, m.span()))

    for m in _RANGE_WHOLE_COL_LOCAL_RE.finditer(formula):
        col = m.group("col").upper()
        ref = CellRef(sheet=None, column=col, row=0, range_kind="whole_column")
        out.append((ref, ref, m.span()))

    for m in _RANGE_WHOLE_ROW_QUOTED_RE.finditer(formula):
        sheet = _sheet_from_quoted_group(m)
        row = int(m.group("row"))
        ref = CellRef(sheet=sheet, column="", row=row, range_kind="whole_row")
        out.append((ref, ref, m.span()))

    for m in _RANGE_WHOLE_ROW_UNQUOTED_RE.finditer(formula):
        sheet = m.group("sheet")
        row = int(m.group("row"))
        ref = CellRef(sheet=sheet, column="", row=row, range_kind="whole_row")
        out.append((ref, ref, m.span()))

    for m in _RANGE_WHOLE_ROW_LOCAL_RE.finditer(formula):
        row = int(m.group("row"))
        ref = CellRef(sheet=None, column="", row=row, range_kind="whole_row")
        out.append((ref, ref, m.span()))

    for m in _RANGE_QUOTED_BOTH_ENDPOINTS_RE.finditer(formula):
        sheet = _sheet_from_quoted_group(m)
        out.append(
            (
                CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
                m.span(),
            )
        )

    for m in _RANGE_UNQUOTED_BOTH_ENDPOINTS_RE.finditer(formula):
        sheet = m.group("sheet")
        out.append(
            (
                CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
                CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
                m.span(),
            )
        )

    for m in _RANGE_QUOTED_RE.finditer(formula):
        sheet = _sheet_from_quoted_group(m)
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
    """Expand an A1 range into individual `(sheet, A1)` dependencies.

    A range whose cell count is at most `max_cells` is enumerated in full
    (row-major). A range of exactly `max_cells` cells is included; one extra
    cell is enough to reject the rectangle.

    Args:
        sheet: Sheet that owns the rectangle.
        start_col: Start column letters (either corner).
        start_row: Start row (either corner).
        end_col: End column letters (either corner).
        end_row: End row (either corner).
        max_cells: Inclusive expansion budget.

    Returns:
        One `(sheet, A1)` pair per cell in the rectangle.

    Raises:
        ValueError: If the rectangle has more than `max_cells` cells.
    """
    c1i = fastpyxl.utils.cell.column_index_from_string(start_col)
    c2i = fastpyxl.utils.cell.column_index_from_string(end_col)
    rlo, rhi = (start_row, end_row) if start_row <= end_row else (end_row, start_row)
    clo, chi = (c1i, c2i) if c1i <= c2i else (c2i, c1i)
    n_cells = (rhi - rlo + 1) * (chi - clo + 1)
    if n_cells > max_cells:
        label = format_range_key(sheet, f"{start_col}{start_row}", f"{end_col}{end_row}")
        raise ValueError(f"Range {label} has {n_cells} cells, exceeding max_cells={max_cells}")

    out: list[tuple[str, str]] = []
    for rr in range(rlo, rhi + 1):
        for cc in range(clo, chi + 1):
            out.append((sheet, f"{fastpyxl.utils.cell.get_column_letter(cc)}{rr}"))
    return out


def expand_range_ref(
    *,
    start: CellRef,
    end: CellRef,
    default_sheet: str,
    max_cells: int,
    sheet_bounds: SheetBounds | None = None,
) -> list[tuple[str, str]]:
    """Expand a parsed range reference, using workbook bounds for whole-column/row forms.

    Rectangular ranges follow `expand_range` and raise `ValueError` when they
    exceed `max_cells`. Whole-column/row shorthands expand to the used-range
    extent regardless of `max_cells`.
    """
    sheet = start.sheet if start.sheet is not None else default_sheet
    bounds = sheet_bounds or {}

    if start.range_kind == "whole_column":
        return expand_whole_column_deps(sheet, start.column, bounds)
    if start.range_kind == "whole_row":
        return expand_whole_row_deps(sheet, start.row, bounds)

    return expand_range(
        sheet=sheet,
        start_col=start.column,
        start_row=start.row,
        end_col=end.column,
        end_row=end.row,
        max_cells=max_cells,
    )


def mask_spans(s: str, spans: list[tuple[int, int]]) -> str:
    """Replace characters in the specified spans with spaces.

    Useful to prevent endpoint tokens from being re-parsed as separate refs.
    """
    if not spans:
        return s
    buf = list(s)
    for a, b in spans:
        for i in range(a, b):
            buf[i] = " "
    return "".join(buf)


def ref_only_function_spans(formula: str) -> list[tuple[int, int]]:
    """Return spans of address/shape-only `ROW`/`COLUMN`/`ROWS`/`COLUMNS` calls.

    Only calls whose arguments contain no nested function calls are included.
    Those arguments are used solely for address/shape (`is_ref_only_arg`), so
    they must not create value dependency edges. Nested calls such as
    `ROW(INDEX(...))` are left unmasked so refs inside the nested expression
    remain visible for dependency extraction.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []
    spans: list[tuple[int, int]] = []
    for _fn, inner, span in _find_function_calls_with_spans(formula, ref_only_function_names()):
        # Nested calls may evaluate values (e.g. INDEX row selector); keep them.
        if "(" in inner:
            continue
        spans.append(span)
    return spans


def mask_ref_only_function_calls(formula: str) -> str:
    """Mask address-only `ROW`/`COLUMN`/`ROWS`/`COLUMNS` calls in `formula`."""
    return mask_spans(formula, ref_only_function_spans(formula))


# Private aliases so the module's historical internal names keep working.
_needs_quoting = needs_quoting
_format_ref = format_cell_key


def normalize_formula(
    formula: str,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> str:
    """Regex-normalize a formula (transitional; not the AST render dialect).

    - Replace same-sheet refs (A1) with sheet-qualified refs (Sheet1!A1)
    - Resolve named ranges to their targets
    - Strip absolute markers ($)
    - Qualify range endpoints

    Returns:
        The regex-normalized formula string. Do not record provenance spans
        against this text and later apply them to `Node.normalized_formula`.
    """
    return normalize_excel_formula(
        formula,
        current_sheet,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )


class FormulaNormalizer:
    """Cached regex formula normalizer (transitional, non-authoritative).

    AST `render_formula` (`A1_ABSOLUTE`) is the `normalized_formula` dialect.
    Use this class for unparseable-cell fallback and for string-based ref
    scans. Do not treat its output as a peer of `Node.normalized_formula`.

    Compared to calling `normalize_formula` repeatedly:

    - Named-range substitution is done in a **single regex pass** over the
      formula string (one compiled alternation pattern for all names), rather
      than one `re.sub` call per name.  This reduces per-call cost from
      O(names) to O(formula_length).
    - Results are **cached** by `(formula, current_sheet)` for the lifetime
      of the normalizer, so repeated calls (common during graph traversal) are
      O(1) dictionary lookups.

    Usage::

        normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
        norm = normalizer.normalize(formula, current_sheet)
    """

    def __init__(
        self,
        named_ranges: dict[str, tuple[str, str]] | None = None,
        named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
    ) -> None:
        self._replacements, self._names_re = build_named_range_replacement_state(
            named_ranges,
            named_range_ranges,
        )

        # Per-instance cache: (formula, current_sheet) -> normalized string
        self._cache: dict[tuple[str, str], str] = {}

    def normalize(self, formula: str, current_sheet: str) -> str:
        """Return the normalized form of *formula*, using the cache when available."""
        if not formula or not formula.startswith("="):
            return formula

        key = (formula, current_sheet)
        cached = self._cache.get(key)
        if cached is not None:
            return cached

        result = self._compute(formula, current_sheet)
        self._cache[key] = result
        return result

    def _compute(self, formula: str, current_sheet: str) -> str:
        """Run the full normalization pipeline (no caching)."""
        return normalize_excel_formula_with_name_state(
            formula,
            current_sheet,
            replacements=self._replacements,
            names_re=self._names_re,
        )


@functools.lru_cache(maxsize=4096)
def split_top_level_if(formula: str) -> tuple[str, str, str] | None:
    """If formula is a top-level IF(...), return (cond, then_expr, else_expr) strings."""
    args = split_top_level_function(formula, "IF")
    if args is None or len(args) != 3:
        return None
    cond, then_expr, else_expr = args
    if not cond or not then_expr:
        return None
    # Excel allows empty else, treat as "" expression (no deps).
    return cond, then_expr, else_expr


def split_top_level_function(formula: str, fn: str) -> list[str] | None:
    """If formula is a top-level FN(...), return the top-level argument strings.

    This is a lightweight splitter that is aware of nested parentheses and quoted
    string literals (") so commas inside those don't split arguments.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    s = formula[1:].lstrip()
    fn_u = fn.upper()
    prefixes = excel_function_call_prefixes(fn_u)
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


@functools.lru_cache(maxsize=4096)
def split_top_level_ifs(formula: str) -> list[str] | None:
    """If formula is a top-level IFS(...), return argument strings."""
    return split_top_level_function(formula, "IFS")


@functools.lru_cache(maxsize=4096)
def split_top_level_choose(formula: str) -> list[str] | None:
    """If formula is a top-level CHOOSE(...), return argument strings."""
    return split_top_level_function(formula, "CHOOSE")


@functools.lru_cache(maxsize=4096)
def split_top_level_switch(formula: str) -> list[str] | None:
    """If formula is a top-level SWITCH(...), return argument strings."""
    return split_top_level_function(formula, "SWITCH")


_CELL_TOKEN_RE = re.compile(
    rf"^(?:{_QUOTED_SHEET_QS}|(?P<us>[A-Za-z][A-Za-z0-9_]*))!\$?(?P<col>[A-Z]{{1,3}})\$?(?P<row>\d+)$"
)
_LOCAL_CELL_TOKEN_RE = re.compile(r"^\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)$")
_NUMBER_TOKEN_RE = re.compile(r"^-?\d+(?:\.\d+)?$")
_INT_TOKEN_RE = re.compile(r"^[+-]?\d+$")
_RANGE_TOKEN_QUOTED_RE = re.compile(
    "^"
    + _QUOTED_SHEET_PREFIX
    + r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)$"
)
_RANGE_TOKEN_UNQUOTED_RE = re.compile(
    r"^(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)$"
)
_RANGE_TOKEN_LOCAL_RE = re.compile(
    r"^\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)$"
)


def _to_node_key(sheet: str, col: str, row: int) -> NodeKey:
    return format_cell_key(sheet, col, row)


def _parse_int_literal(s: str) -> int | None:
    s = s.strip()
    if not _INT_TOKEN_RE.match(s):
        return None
    try:
        return int(s)
    except ValueError:
        return None


_WARNED_CACHED_DYNAMIC = False


def _warn_cached_dynamic_once() -> None:
    global _WARNED_CACHED_DYNAMIC
    if _WARNED_CACHED_DYNAMIC:
        return
    _WARNED_CACHED_DYNAMIC = True
    warnings.warn(
        "Resolved OFFSET/INDIRECT references from cached workbook values; these "
        "dependencies are fixed at graph-build time. Changing an input that shifts a "
        "resolution target outside the graph makes the graph uncomputable. Pass "
        "`dynamic_refs` to resolve over an input domain instead.",
        UserWarning,
        stacklevel=2,
    )


def _parse_string_literal(s: str) -> str | None:
    s = s.strip()
    if len(s) < 2 or s[0] != '"' or s[-1] != '"':
        return None
    # Excel escapes quotes as doubled quotes within string literals.
    return s[1:-1].replace('""', '"')


def _eval_row_expr(
    expr: str,
    *,
    current_sheet: str,
    current_cell_a1: str | None,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]] | None,
) -> int | None:
    s = expr.strip()
    if not s:
        return None
    # Reject unsupported operators
    if "*" in s or "/" in s or "^" in s:
        return None

    def eval_term(term: str) -> int | None:
        term = term.strip()
        if not term:
            return None
        lit = _parse_int_literal(term)
        if lit is not None:
            return lit
        m = re.match(r"^ROW\s*\((?P<inner>.*)\)$", term, flags=re.IGNORECASE)
        if m:
            inner = m.group("inner").strip()
            if inner == "":
                if current_cell_a1 is None:
                    return None
                _col, row = fastpyxl.utils.cell.coordinate_from_string(current_cell_a1)
                return int(row)
            parsed = _parse_ref_or_range_token(
                inner,
                current_sheet=current_sheet,
                named_ranges=named_ranges,
                named_range_ranges=named_range_ranges,
            )
            if parsed is None:
                return None
            start_ref, end_ref = parsed
            if start_ref.sheet is None:
                start_ref = CellRef(sheet=current_sheet, column=start_ref.column, row=start_ref.row)
            if end_ref.sheet is None:
                end_ref = CellRef(sheet=current_sheet, column=end_ref.column, row=end_ref.row)
            if (
                (start_ref.sheet != end_ref.sheet)
                or (start_ref.column != end_ref.column)
                or (start_ref.row != end_ref.row)
            ):
                return None
            return start_ref.row
        return None

    parts = re.split(r"([+-])", s)
    if not parts:
        return None
    acc = eval_term(parts[0])
    if acc is None:
        return None
    idx = 1
    while idx < len(parts):
        op = parts[idx]
        term = parts[idx + 1] if idx + 1 < len(parts) else ""
        val = eval_term(term)
        if val is None:
            return None
        if op == "+":
            acc += val
        elif op == "-":
            acc -= val
        else:
            return None
        idx += 2
    return acc


def _resolve_numeric_token(
    token: str,
    *,
    current_sheet: str,
    current_cell_a1: str | None,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]] | None,
    value_resolver: Callable[[str, str], object] | None,
) -> int | None:
    lit = _parse_int_literal(token)
    if lit is not None:
        return lit

    row_expr = _eval_row_expr(
        token,
        current_sheet=current_sheet,
        current_cell_a1=current_cell_a1,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )
    if row_expr is not None:
        return row_expr

    parsed = _parse_ref_or_range_token(
        token,
        current_sheet=current_sheet,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )
    if parsed is None:
        return None
    start_ref, end_ref = parsed
    if (
        (start_ref.sheet != end_ref.sheet)
        or (start_ref.column != end_ref.column)
        or (start_ref.row != end_ref.row)
    ):
        return None
    sheet = start_ref.sheet if start_ref.sheet is not None else current_sheet
    if value_resolver is None:
        return None
    raw = value_resolver(sheet, f"{start_ref.column}{start_ref.row}")
    if isinstance(raw, bool):
        return None
    if isinstance(raw, int):
        _warn_cached_dynamic_once()
        return raw
    if isinstance(raw, float) and raw.is_integer():
        _warn_cached_dynamic_once()
        return int(raw)
    return None


def _parse_index_call(
    token: str,
    *,
    current_sheet: str,
    current_cell_a1: str | None,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]] | None,
) -> CellRef | None:
    s = token.strip()
    if not s.upper().startswith("INDEX("):
        return None
    if not s.endswith(")"):
        return None
    inner = s[s.find("(") + 1 : -1]
    args = _split_function_args(inner)
    if args is None or len(args) < 2:
        return None
    base_arg = args[0]
    row_arg = args[1]
    col_arg = args[2] if len(args) >= 3 else "1"

    base = _parse_ref_or_range_token(
        base_arg,
        current_sheet=current_sheet,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )
    if base is None:
        return None
    base_start, base_end = base
    sheet = base_start.sheet or current_sheet

    row_num = _resolve_numeric_token(
        row_arg,
        current_sheet=current_sheet,
        current_cell_a1=current_cell_a1,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
        value_resolver=None,
    )
    if row_num is None:
        return None
    col_num = _parse_int_literal(col_arg)
    if col_num is None:
        return None

    start_col = fastpyxl.utils.cell.column_index_from_string(base_start.column)
    end_col = fastpyxl.utils.cell.column_index_from_string(base_end.column)
    start_row = base_start.row
    end_row = base_end.row
    min_col = min(start_col, end_col)
    max_col = max(start_col, end_col)
    min_row = min(start_row, end_row)
    max_row = max(start_row, end_row)

    if row_num < 1 or col_num < 1:
        return None
    target_row = min_row + row_num - 1
    target_col = min_col + col_num - 1
    if target_row < min_row:
        target_row = min_row
    elif target_row > max_row:
        target_row = max_row
    if target_col < min_col:
        target_col = min_col
    elif target_col > max_col:
        target_col = max_col

    return CellRef(
        sheet=sheet,
        column=fastpyxl.utils.cell.get_column_letter(target_col),
        row=target_row,
    )


def _parse_ref_or_range_token(
    token: str,
    *,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> tuple[CellRef, CellRef] | None:
    tok = token.strip()
    if not tok:
        return None

    # Named range (single-cell)
    if tok in named_ranges:
        sheet, addr = named_ranges[tok]
        m = re.match(r"^(?P<col>[A-Z]{1,3})(?P<row>\d+)$", addr)
        if not m:
            return None
        ref = CellRef(sheet=sheet, column=m.group("col"), row=int(m.group("row")))
        return ref, ref

    # Named range (range)
    if named_range_ranges is not None and tok in named_range_ranges:
        sheet, start_a1, end_a1 = named_range_ranges[tok]
        m1 = re.match(r"^(?P<col>[A-Z]{1,3})(?P<row>\d+)$", start_a1)
        m2 = re.match(r"^(?P<col>[A-Z]{1,3})(?P<row>\d+)$", end_a1)
        if not m1 or not m2:
            return None
        return (
            CellRef(sheet=sheet, column=m1.group("col"), row=int(m1.group("row"))),
            CellRef(sheet=sheet, column=m2.group("col"), row=int(m2.group("row"))),
        )

    # Quoted sheet range
    m = _RANGE_TOKEN_QUOTED_RE.match(tok)
    if m:
        sheet = _sheet_from_quoted_group(m)
        return (
            CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
            CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
        )

    # Unquoted sheet range
    m = _RANGE_TOKEN_UNQUOTED_RE.match(tok)
    if m:
        sheet = m.group("sheet")
        return (
            CellRef(sheet=sheet, column=m.group("c1"), row=int(m.group("r1"))),
            CellRef(sheet=sheet, column=m.group("c2"), row=int(m.group("r2"))),
        )

    # Local range
    m = _RANGE_TOKEN_LOCAL_RE.match(tok)
    if m:
        return (
            CellRef(sheet=None, column=m.group("c1"), row=int(m.group("r1"))),
            CellRef(sheet=None, column=m.group("c2"), row=int(m.group("r2"))),
        )

    # Sheet-qualified cell
    m = _CELL_TOKEN_RE.match(tok)
    if m:
        sheet = _sheet_from_cell_match(m)
        ref = CellRef(sheet=sheet, column=m.group("col"), row=int(m.group("row")))
        return ref, ref

    # Local cell
    m = _LOCAL_CELL_TOKEN_RE.match(tok)
    if m:
        col = m.group("col")
        if col in _FUNC_LIKE:
            return None
        ref = CellRef(sheet=None, column=col, row=int(m.group("row")))
        return ref, ref

    return None


def _split_function_args(inner: str) -> list[str] | None:
    return _split_top_level_args(inner)


@functools.lru_cache(maxsize=4096)
def _find_function_calls_with_spans(
    formula: str, fn_names: frozenset[str]
) -> list[tuple[str, str, tuple[int, int]]]:
    s = formula
    out: list[tuple[str, str, tuple[int, int]]] = []
    i = 0
    in_str = False
    n = len(s)
    while i < n:
        ch = s[i]
        if ch == '"':
            in_str = not in_str
            i += 1
            continue
        if in_str:
            i += 1
            continue
        if ch.isalpha() or ch == "_":
            start = i
            j = i + 1
            while j < n and (s[j].isalnum() or s[j] in "_."):
                j += 1
            token = s[start:j]
            token_u = token.upper()
            fn = None
            if token_u in fn_names:
                fn = token_u
            else:
                for name in fn_names:
                    if token_u.endswith(f".{name}"):
                        fn = name
                        break
            if fn is not None:
                k = j
                while k < n and s[k].isspace():
                    k += 1
                if k < n and s[k] == "(":
                    depth = 0
                    in_call_str = False
                    m = k
                    while m < n:
                        ch2 = s[m]
                        if ch2 == '"':
                            in_call_str = not in_call_str
                            m += 1
                            continue
                        if in_call_str:
                            m += 1
                            continue
                        if ch2 == "(":
                            depth += 1
                        elif ch2 == ")":
                            depth -= 1
                            if depth == 0:
                                inner = s[k + 1 : m]
                                out.append((fn, inner, (start, m + 1)))
                                i = m + 1
                                break
                        m += 1
                    else:
                        i = j
                    continue
            i = j
            continue
        i += 1
    return out


def _parse_offset_call(
    args: list[str],
    *,
    current_sheet: str,
    current_cell_a1: str | None,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]] | None,
    value_resolver: Callable[[str, str], object] | None,
) -> tuple[CellRef, CellRef]:
    if len(args) < 3 or len(args) > 5:
        raise ValueError("OFFSET expects 3 to 5 arguments")
    base_arg, rows_arg, cols_arg = args[0], args[1], args[2]
    height_arg = args[3] if len(args) >= 4 else None
    width_arg = args[4] if len(args) >= 5 else None

    base = _parse_ref_or_range_token(
        base_arg,
        current_sheet=current_sheet,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )
    if base is None:
        index_ref = _parse_index_call(
            base_arg,
            current_sheet=current_sheet,
            current_cell_a1=current_cell_a1,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
        if index_ref is None:
            raise ValueError("OFFSET base must be a cell or range reference")
        base_start = index_ref
        base_end = index_ref
    else:
        base_start, base_end = base
    sheet = base_start.sheet or current_sheet

    base_start_col = fastpyxl.utils.cell.column_index_from_string(base_start.column)
    base_end_col = fastpyxl.utils.cell.column_index_from_string(base_end.column)
    base_start_row = base_start.row
    base_end_row = base_end.row
    base_min_col = min(base_start_col, base_end_col)
    base_max_col = max(base_start_col, base_end_col)
    base_min_row = min(base_start_row, base_end_row)
    base_max_row = max(base_start_row, base_end_row)

    rows = _resolve_numeric_token(
        rows_arg,
        current_sheet=current_sheet,
        current_cell_a1=current_cell_a1,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
        value_resolver=value_resolver,
    )
    cols = _resolve_numeric_token(
        cols_arg,
        current_sheet=current_sheet,
        current_cell_a1=current_cell_a1,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
        value_resolver=value_resolver,
    )
    if rows is None or cols is None:
        raise ValueError("OFFSET rows/cols must be integer literals or cached numeric refs")

    base_height = base_max_row - base_min_row + 1
    base_width = base_max_col - base_min_col + 1

    height = (
        base_height
        if height_arg is None or height_arg == ""
        else _resolve_numeric_token(
            height_arg,
            current_sheet=current_sheet,
            current_cell_a1=current_cell_a1,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
            value_resolver=value_resolver,
        )
    )
    width = (
        base_width
        if width_arg is None or width_arg == ""
        else _resolve_numeric_token(
            width_arg,
            current_sheet=current_sheet,
            current_cell_a1=current_cell_a1,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
            value_resolver=value_resolver,
        )
    )
    if height is None or width is None:
        raise ValueError(
            "OFFSET height/width must be integer literals or cached numeric refs when provided"
        )
    if height <= 0 or width <= 0:
        raise ValueError("OFFSET height/width must be positive")

    target_start_row = base_min_row + rows
    target_start_col = base_min_col + cols
    if target_start_row <= 0 or target_start_col <= 0:
        raise ValueError("OFFSET target reference is out of bounds")

    target_end_row = target_start_row + height - 1
    target_end_col = target_start_col + width - 1
    if target_end_row <= 0 or target_end_col <= 0:
        raise ValueError("OFFSET target reference is out of bounds")

    start_ref = CellRef(
        sheet=sheet,
        column=fastpyxl.utils.cell.get_column_letter(target_start_col),
        row=target_start_row,
    )
    end_ref = CellRef(
        sheet=sheet,
        column=fastpyxl.utils.cell.get_column_letter(target_end_col),
        row=target_end_row,
    )
    return start_ref, end_ref


def _parse_indirect_call(
    args: list[str],
    *,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]] | None,
) -> tuple[CellRef, CellRef]:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("INDIRECT expects 1 or 2 arguments")
    text_arg = args[0]
    literal = _parse_string_literal(text_arg)
    if literal is None:
        raise ValueError("INDIRECT text must be a string literal")

    if len(args) == 2:
        a1_arg = args[1].strip()
        if a1_arg.upper() in {"TRUE", "1"}:
            pass
        elif a1_arg.upper() in {"FALSE", "0"}:
            raise ValueError("INDIRECT R1C1 style is not supported")
        else:
            raise ValueError("INDIRECT A1/R1C1 flag must be a literal TRUE/FALSE or 1/0")

    parsed = _parse_ref_or_range_token(
        literal,
        current_sheet=current_sheet,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )
    if parsed is None:
        raise ValueError("INDIRECT text must resolve to a cell or range reference")
    start_ref, end_ref = parsed
    if start_ref.sheet is None:
        start_ref = CellRef(sheet=current_sheet, column=start_ref.column, row=start_ref.row)
    if end_ref.sheet is None:
        end_ref = CellRef(sheet=current_sheet, column=end_ref.column, row=end_ref.row)
    return start_ref, end_ref


def parse_dynamic_range_refs_with_spans(
    formula: str,
    *,
    current_sheet: str,
    current_cell_a1: str | None = None,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
    normalizer: FormulaNormalizer | None = None,
    value_resolver: Callable[[str, str], object] | None = None,
) -> list[tuple[CellRef, CellRef, tuple[int, int], list[CellRef]]]:
    """Extract dynamic range references (OFFSET/INDIRECT).

    Returns:
        Tuples of (start, end, span, arg_refs).

    Raises:
        ValueError: Non-literal offsets or non-literal INDIRECT references.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []
    if named_ranges is None:
        named_ranges = {}
    if normalizer is None:
        normalizer = FormulaNormalizer(named_ranges, named_range_ranges)

    calls = _find_function_calls_with_spans(formula, frozenset({"OFFSET", "INDIRECT"}))
    out: list[tuple[CellRef, CellRef, tuple[int, int], list[CellRef]]] = []
    for fn, inner, span in calls:
        args = _split_function_args(inner)
        if args is None:
            raise ValueError(f"{fn} has unbalanced parentheses or strings")
        if fn == "OFFSET":
            arg_refs: list[CellRef] = []
            for arg in args[1:]:
                normalized = normalizer.normalize(
                    "=" + arg,
                    current_sheet,
                )
                value_expr = mask_ref_only_function_calls(normalized)
                for ref in parse_standalone_cell_refs(value_expr):
                    sheet = ref.sheet if ref.sheet is not None else current_sheet
                    arg_refs.append(
                        CellRef(
                            sheet=sheet,
                            column=ref.column,
                            row=ref.row,
                        )
                    )
            start_ref, end_ref = _parse_offset_call(
                args,
                current_sheet=current_sheet,
                current_cell_a1=current_cell_a1,
                named_ranges=named_ranges,
                named_range_ranges=named_range_ranges,
                value_resolver=value_resolver,
            )
        else:
            start_ref, end_ref = _parse_indirect_call(
                args,
                current_sheet=current_sheet,
                named_ranges=named_ranges,
                named_range_ranges=named_range_ranges,
            )
            arg_refs = []
        out.append((start_ref, end_ref, span, arg_refs))
    return out


def parse_guard_expr(
    expr: str,
    *,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    allow_ranges: bool = False,
) -> GuardExpr | None:
    """Parse a minimal boolean expression suitable for IF(...) conditions.

    Args:
        expr: The condition text (no leading `=`).
        current_sheet: Sheet used to qualify unqualified references.
        named_ranges: Single-cell named ranges available to the formula.
        allow_ranges: Parse multi-cell range operands into `RangeRef` template
            nodes, for array-context conditions evaluated element-wise. Ranges
            are still rejected inside `AND`/`OR`, which collapse an array to a
            single boolean rather than evaluating it element-wise.

    Returns:
        The parsed guard, or `None` when the syntax is unsupported.
        Successful results are passed through `intern_guard`.
    """
    parsed = _parse_guard_expr(
        expr,
        current_sheet=current_sheet,
        named_ranges=named_ranges,
        allow_ranges=allow_ranges,
    )
    return None if parsed is None else intern_guard(parsed)


def _parse_guard_expr(
    expr: str,
    *,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    allow_ranges: bool = False,
) -> GuardExpr | None:
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
            operand = parse_guard_expr(
                parts[0],
                current_sheet=current_sheet,
                named_ranges=named_ranges,
                allow_ranges=allow_ranges,
            )
            return None if operand is None else Not(operand)
        if fn in {"AND", "OR"}:
            if len(parts) < 1:
                return None
            ops: list[GuardExpr] = []
            for p in parts:
                # `allow_ranges=False`: AND/OR reduce an array operand to one
                # boolean, so their operands are not element-wise templates.
                ge = parse_guard_expr(p, current_sheet=current_sheet, named_ranges=named_ranges)
                if ge is None:
                    return None
                ops.append(ge)
            return And(tuple(ops)) if fn == "AND" else Or(tuple(ops))

    # Comparison: left op right
    for op in ("<>", "<=", ">=", "=", "<", ">"):
        if op in s:
            left_s, right_s = (p.strip() for p in s.split(op, 1))
            left = _parse_guard_atom(
                left_s,
                current_sheet=current_sheet,
                named_ranges=named_ranges,
                allow_ranges=allow_ranges,
            )
            right = _parse_guard_atom(
                right_s,
                current_sheet=current_sheet,
                named_ranges=named_ranges,
                allow_ranges=allow_ranges,
            )
            if left is None or right is None:
                return None
            return Compare(left=left, op=op, right=right)

    # Atom (truthy cell ref or literal)
    return _parse_guard_atom(
        s, current_sheet=current_sheet, named_ranges=named_ranges, allow_ranges=allow_ranges
    )


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
    allow_ranges: bool = False,
) -> GuardExpr | None:
    if not s:
        return None

    # Multi-cell range (array-context conditions only)
    if allow_ranges:
        range_ref = _parse_guard_range_atom(s, current_sheet=current_sheet)
        if range_ref is not None:
            return range_ref

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
        sheet = _sheet_from_cell_match(m)
        return GuardCellRef(_to_node_key(str(sheet), m.group("col"), int(m.group("row"))))

    # Local cell
    m = _LOCAL_CELL_TOKEN_RE.match(s)
    if m:
        col = m.group("col")
        if col in _FUNC_LIKE:
            return None
        return GuardCellRef(_to_node_key(current_sheet, col, int(m.group("row"))))

    return None


def _parse_guard_range_atom(s: str, *, current_sheet: str) -> RangeRef | None:
    """Parse a bare multi-cell range token into a `RangeRef` guard template."""
    for regex, sheet_of in (
        (_RANGE_TOKEN_QUOTED_RE, _sheet_from_quoted_group),
        (_RANGE_TOKEN_UNQUOTED_RE, lambda m: m.group("sheet")),
        (_RANGE_TOKEN_LOCAL_RE, lambda _m: current_sheet),
    ):
        m = regex.match(s)
        if m is None:
            continue
        sheet = sheet_of(m)
        key = _address_keys.parse_node_key(
            _address_keys.format_range_key(
                sheet,
                f"{m.group('c1')}{m.group('r1')}",
                f"{m.group('c2')}{m.group('r2')}",
            )
        )
        if not isinstance(key, _address_keys.RangeKey):
            # A 1x1 "range" is a scalar; leave it to the cell atom parser.
            return None
        return RangeRef(str(key))
    return None


# Functions whose result at each array position depends only on the operands at
# that same position. Ranges nested anywhere else (aggregates such as `SUM`,
# lookups, `OFFSET`/`INDIRECT`/`INDEX`) are not element-aligned with an
# enclosing array condition, so their cells stay unguarded.
_ELEMENT_WISE_FN_NAMES = frozenset(
    {
        "ABS",
        "CEILING",
        "EXP",
        "FLOOR",
        "IF",
        "IFERROR",
        "IFNA",
        "IFS",
        "INT",
        "LN",
        "LOG",
        "LOG10",
        "MOD",
        "MROUND",
        "N",
        "NOT",
        "POWER",
        "ROUND",
        "ROUNDDOWN",
        "ROUNDUP",
        "SIGN",
        "SQRT",
        "SWITCH",
        "T",
        "TRUNC",
        "VALUE",
    }
)

_FUNCTION_CALL_HEAD_RE = re.compile(r"(?P<name>[A-Za-z_][A-Za-z0-9_.]*)\s*\(")


@functools.lru_cache(maxsize=4096)
def _find_all_function_call_spans(formula: str) -> tuple[tuple[str, tuple[int, int]], ...]:
    """Return `(function_name, span)` for every function call in `formula`."""
    out: list[tuple[str, tuple[int, int]]] = []
    for m in _FUNCTION_CALL_HEAD_RE.finditer(formula):
        open_idx = m.end() - 1
        depth = 0
        in_str = False
        i = open_idx
        while i < len(formula):
            ch = formula[i]
            if ch == '"':
                in_str = not in_str
            elif not in_str:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        name = m.group("name").upper().rsplit(".", 1)[-1]
                        out.append((name, (m.start(), i + 1)))
                        break
            i += 1
    return tuple(out)


def element_aligned_range_cells(
    expr: str,
    *,
    current_sheet: str,
    shape: tuple[int, int],
    max_cells: int,
) -> dict[tuple[str, str], tuple[int, int]]:
    """Map cells of `expr`'s element-aligned ranges to their `(row, col)` offsets.

    A range is element-aligned with an enclosing array condition when it has the
    same shape and every function call enclosing it evaluates element-wise (see
    `_ELEMENT_WISE_FN_NAMES`). Cells that would take two different offsets (from
    overlapping ranges) are dropped, since no single element guard describes them.

    Args:
        expr: Branch expression text (with or without a leading `=`).
        current_sheet: Sheet used to qualify unqualified references.
        shape: `(n_rows, n_cols)` of the condition's range.
        max_cells: Range-expansion budget; larger ranges are skipped because
            their dependencies are not expanded per cell either.

    Returns:
        A `(sheet, a1) -> (row_offset, col_offset)` mapping, possibly empty.
    """
    f = expr if expr.startswith("=") else "=" + expr
    call_spans = _find_all_function_call_spans(f)

    offsets: dict[tuple[str, str], tuple[int, int]] = {}
    ambiguous: set[tuple[str, str]] = set()
    for start, end, span in parse_range_refs_with_spans(f):
        if start.range_kind != "cell":
            continue
        if any(
            name not in _ELEMENT_WISE_FN_NAMES
            for name, call_span in call_spans
            if call_span[0] <= span[0] and span[1] <= call_span[1]
        ):
            continue

        sheet = start.sheet if start.sheet is not None else current_sheet
        c1 = fastpyxl.utils.cell.column_index_from_string(start.column)
        c2 = fastpyxl.utils.cell.column_index_from_string(end.column)
        col_lo, col_hi = (c1, c2) if c1 <= c2 else (c2, c1)
        row_lo, row_hi = (start.row, end.row) if start.row <= end.row else (end.row, start.row)
        n_rows = row_hi - row_lo + 1
        n_cols = col_hi - col_lo + 1
        if (n_rows, n_cols) != shape or n_rows * n_cols > max_cells:
            continue

        for row_offset in range(n_rows):
            for col_offset in range(n_cols):
                column = fastpyxl.utils.cell.get_column_letter(col_lo + col_offset)
                cell = (sheet, f"{column}{row_lo + row_offset}")
                existing = offsets.get(cell)
                if existing is not None and existing != (row_offset, col_offset):
                    ambiguous.add(cell)
                offsets[cell] = (row_offset, col_offset)

    for cell in ambiguous:
        offsets.pop(cell, None)
    return offsets
