"""Transitional regex Excel formula normalization for sheet-qualified A1 text.

The formula-text dialect of record is AST `render_formula` (`A1_ABSOLUTE`).
Regex normalization is **not** a peer of that dialect: it does not match AST
canonicalizations (unary `+`, redundant parens, number spelling, compact
whitespace, lowercase refs). Graph extraction uses it only as a fallback when
the AST parser cannot handle a cell, and as a helper for string-based ref
scans of already-rendered text.

`expand_defined_names` is a parse preprocessor (keeps `$` markers) used by
`parse_preserving_axes`; it is not the `normalized_formula` dialect.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from excel_grapher.core.address_keys import (
    format_cell_key,
    format_range_key,
    parse_address,
    quote_sheet_if_needed,
    quoted_sheet_prefix_regex,
    unescape_formula_sheet_name,
)

_QUOTED_SHEET_PREFIX = quoted_sheet_prefix_regex()

_FUNC_LIKE = frozenset({"IF", "OR", "AND", "NOT", "SUM", "MAX", "MIN", "AVG"})
_STRING_LITERAL_RE = re.compile(r'"(?:[^"]|"")*"')


def _mask_string_literals(formula: str) -> tuple[str, list[str]]:
    """Replace Excel string literals with placeholders so normalization skips them."""
    literals: list[str] = []

    def _stash(match: re.Match[str]) -> str:
        literals.append(match.group(0))
        return f"__XL_STR_{len(literals) - 1}__"

    return _STRING_LITERAL_RE.sub(_stash, formula), literals


def _unmask_string_literals(formula: str, literals: list[str]) -> str:
    """Restore stashed string literals after normalization."""
    for index, literal in enumerate(literals):
        formula = formula.replace(f"__XL_STR_{index}__", literal)
    return formula


@dataclass(frozen=True, slots=True)
class NamedRangeReplacementState:
    """Compiled defined-name substitutions for regex normalize and parse.

    `replacements` are sheet-qualified A1 without `$` (regex normalizer).
    `absolute_replacements` keep `$` on both axes so `parse_preserving_axes`
    treats defined names as absolute.
    """

    replacements: dict[str, str]
    names_re: re.Pattern[str] | None
    absolute_replacements: dict[str, str]


def build_named_range_replacement_state(
    named_ranges: dict[str, tuple[str, str]] | None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None,
) -> NamedRangeReplacementState:
    """Build replacement strings and a single alternation regex for defined names."""
    named_ranges = named_ranges or {}
    named_range_ranges = named_range_ranges or {}

    replacements: dict[str, str] = {}

    for name, (sheet, addr) in named_ranges.items():
        col_match = re.match(r"^([A-Z]{1,3})(\d+)$", addr)
        if col_match:
            replacements[name] = format_cell_key(sheet, col_match.group(1), int(col_match.group(2)))

    for name, (sheet, start_a1, end_a1) in named_range_ranges.items():
        m_start = re.match(r"^([A-Z]{1,3})(\d+)$", start_a1)
        m_end = re.match(r"^([A-Z]{1,3})(\d+)$", end_a1)
        if m_start and m_end:
            start_cell = f"{m_start.group(1)}{int(m_start.group(2))}"
            end_cell = f"{m_end.group(1)}{int(m_end.group(2))}"
            replacements[name] = format_range_key(sheet, start_cell, end_cell)

    if not replacements:
        return NamedRangeReplacementState({}, None, {})

    names = sorted(replacements, key=len, reverse=True)
    alt = "|".join(re.escape(n) for n in names)
    names_re = re.compile(rf"\b(?:{alt})\b(?!\s*!)")
    absolute_replacements = {
        name: _absolutize_defined_name_target(target) for name, target in replacements.items()
    }
    return NamedRangeReplacementState(replacements, names_re, absolute_replacements)


def _apply_named_range_replacements(
    formula: str,
    replacements: dict[str, str],
    names_re: re.Pattern[str] | None,
) -> str:
    if names_re is None:
        return formula

    def replace_name(m: re.Match[str]) -> str:
        return replacements.get(m.group(0), m.group(0))

    return names_re.sub(replace_name, formula)


def _format_whole_column_ref(sheet: str, column: str) -> str:
    col = column.upper()
    return f"{quote_sheet_if_needed(sheet)}!{col}:{col}"


def _format_whole_row_ref(sheet: str, row: int) -> str:
    return f"{quote_sheet_if_needed(sheet)}!{row}:{row}"


def _normalize_whole_column_row_shorthand(formula: str, current_sheet: str) -> str:
    """Strip ``$`` and sheet-qualify whole-column/row shorthand without expanding bounds."""
    result = formula

    def quoted_whole_col(m: re.Match[str]) -> str:
        sheet = unescape_formula_sheet_name(m.group("sheet"))
        return _format_whole_column_ref(sheet, m.group("col"))

    result = re.sub(
        _QUOTED_SHEET_PREFIX + r"\$?(?P<col>[A-Z]{1,3})\s*:\s*\$?(?P=col)\b",
        quoted_whole_col,
        result,
        flags=re.IGNORECASE,
    )

    def unquoted_whole_col(m: re.Match[str]) -> str:
        return _format_whole_column_ref(m.group("sheet"), m.group("col"))

    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<col>[A-Z]{1,3})\s*:\s*\$?(?P=col)\b",
        unquoted_whole_col,
        result,
        flags=re.IGNORECASE,
    )

    def local_whole_col(m: re.Match[str]) -> str:
        col = m.group("col")
        if col in _FUNC_LIKE:
            return m.group(0)
        return _format_whole_column_ref(current_sheet, col)

    result = re.sub(
        r"(?<![!A-Za-z0-9_'])(?<!\$)\$?(?P<col>[A-Z]{1,3})\s*:\s*\$?(?P=col)\b(?![A-Za-z0-9_])",
        local_whole_col,
        result,
        flags=re.IGNORECASE,
    )

    def quoted_whole_row(m: re.Match[str]) -> str:
        sheet = unescape_formula_sheet_name(m.group("sheet"))
        return _format_whole_row_ref(sheet, int(m.group("row")))

    result = re.sub(
        _QUOTED_SHEET_PREFIX + r"\$?(?P<row>\d+)\s*:\s*\$?(?P=row)\b",
        quoted_whole_row,
        result,
    )

    def unquoted_whole_row(m: re.Match[str]) -> str:
        return _format_whole_row_ref(m.group("sheet"), int(m.group("row")))

    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<row>\d+)\s*:\s*\$?(?P=row)\b",
        unquoted_whole_row,
        result,
    )

    def local_whole_row(m: re.Match[str]) -> str:
        return _format_whole_row_ref(current_sheet, int(m.group("row")))

    result = re.sub(
        r"(?<![!A-Za-z0-9_'])(?<!\$)(?P<row>\d+)\s*:\s*\$?(?P=row)\b(?![A-Za-z0-9_])",
        local_whole_row,
        result,
    )

    return result


def expand_whole_column_row_for_parse(
    formula: str,
    bounds: dict[str, tuple[int, int]],
) -> str:
    """Expand whole-column/row shorthand to bounded A1 ranges for parse-only paths.

    Used when a workbook is available outside graph build (e.g. defined-name OFFSET
    resolution). Strips ``$`` markers before matching.
    """
    from excel_grapher.core.range_shorthand import (
        whole_column_to_bounded_a1,
        whole_row_to_bounded_a1,
    )

    s = formula.replace("$", "")
    for sheet in bounds:
        quoted = re.escape(quote_sheet_if_needed(sheet))
        bare = re.escape(sheet)

        def repl_col(m: re.Match[str], *, _sheet: str = sheet) -> str:
            start, end = whole_column_to_bounded_a1(_sheet, m.group(1), bounds)
            return f"{start}:{end}"

        for prefix in (quoted, bare):
            s = re.sub(
                prefix + r"!\s*([A-Z]+)\s*:\s*\1\b",
                repl_col,
                s,
                flags=re.IGNORECASE,
            )

        def repl_row(m: re.Match[str], *, _sheet: str = sheet) -> str:
            row = int(m.group(1))
            start, end = whole_row_to_bounded_a1(_sheet, row, bounds)
            return f"{start}:{end}"

        for prefix in (quoted, bare):
            s = re.sub(
                prefix + r"!\s*(\d+)\s*:\s*\1\b",
                repl_row,
                s,
            )
    return s


def _normalize_excel_formula_base(formula: str, current_sheet: str) -> str:
    """Strip $ markers, qualify ranges and cells, without defined-name substitution."""
    result = _normalize_whole_column_row_shorthand(formula, current_sheet)

    def replace_quoted_both_end_range(m: re.Match[str]) -> str:
        sheet1 = unescape_formula_sheet_name(m.group("sheet1"))
        sheet2 = unescape_formula_sheet_name(m.group("sheet2"))
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        start_cell = f"{c1}{int(r1)}"
        end_cell = f"{c2}{int(r2)}"
        if sheet1 == sheet2:
            return format_range_key(sheet1, start_cell, end_cell)
        return f"{format_cell_key(sheet1, c1, int(r1))}:{format_cell_key(sheet2, c2, int(r2))}"

    result = re.sub(
        _QUOTED_SHEET_PREFIX.replace("?P<sheet>", "?P<sheet1>")
        + r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*"
        + _QUOTED_SHEET_PREFIX.replace("?P<sheet>", "?P<sheet2>")
        + r"\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_quoted_both_end_range,
        result,
    )

    def replace_unquoted_both_end_range(m: re.Match[str]) -> str:
        sheet1, sheet2 = m.group("sheet1"), m.group("sheet2")
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        start_cell = f"{c1}{int(r1)}"
        end_cell = f"{c2}{int(r2)}"
        if sheet1 == sheet2:
            return format_range_key(sheet1, start_cell, end_cell)
        return f"{sheet1}!{start_cell}:{sheet2}!{end_cell}"

    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet1>[A-Za-z][A-Za-z0-9_]*)!"
        r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*"
        r"(?P<sheet2>[A-Za-z][A-Za-z0-9_]*)!"
        r"\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_unquoted_both_end_range,
        result,
    )

    def replace_quoted_range(m: re.Match[str]) -> str:
        sheet = unescape_formula_sheet_name(m.group("sheet"))
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        return format_range_key(sheet, f"{c1}{int(r1)}", f"{c2}{int(r2)}")

    result = re.sub(
        _QUOTED_SHEET_PREFIX
        + r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_quoted_range,
        result,
    )

    def replace_unquoted_range(m: re.Match[str]) -> str:
        sheet = m.group("sheet")
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        return format_range_key(sheet, f"{c1}{int(r1)}", f"{c2}{int(r2)}")

    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_unquoted_range,
        result,
    )

    def replace_local_range(m: re.Match[str]) -> str:
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        return format_range_key(current_sheet, f"{c1}{int(r1)}", f"{c2}{int(r2)}")

    result = re.sub(
        r"(?<![!A-Za-z0-9_])(?<!\$)\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)(?![A-Za-z0-9_])",
        replace_local_range,
        result,
    )

    def replace_quoted_cell(m: re.Match[str]) -> str:
        sheet = unescape_formula_sheet_name(m.group("sheet"))
        col, row = m.group("col"), m.group("row")
        return format_cell_key(sheet, col, int(row))

    result = re.sub(
        _QUOTED_SHEET_PREFIX + r"\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)",
        replace_quoted_cell,
        result,
    )

    def replace_unquoted_cell(m: re.Match[str]) -> str:
        sheet, col, row = m.group("sheet"), m.group("col"), m.group("row")
        return f"{sheet}!{col}{row}"

    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)",
        replace_unquoted_cell,
        result,
    )

    def replace_local_cell(m: re.Match[str]) -> str:
        col, row = m.group("col"), m.group("row")
        if col in _FUNC_LIKE:
            return m.group(0)
        return format_cell_key(current_sheet, col, int(row))

    result = re.sub(
        r"(?<![!A-Za-z0-9_:])(?<!\$)\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)(?![A-Za-z0-9_!'])",
        replace_local_cell,
        result,
    )

    return _collapse_same_sheet_range_prefixes(result)


def _collapse_same_sheet_range_prefixes(formula: str) -> str:
    """Collapse same-sheet both-end ranges to single-prefix form.

    Earlier passes already emit single-prefix ranges and exclude `:` from local
    cell lookbehinds. This pass is defense-in-depth for residual both-end forms
    (for example after defined-name substitution) and leaves cross-sheet ranges
    unchanged.
    """

    def replace_quoted(m: re.Match[str]) -> str:
        sheet1 = unescape_formula_sheet_name(m.group("sheet1"))
        sheet2 = unescape_formula_sheet_name(m.group("sheet2"))
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        if sheet1 != sheet2:
            return m.group(0)
        return format_range_key(sheet1, f"{c1}{int(r1)}", f"{c2}{int(r2)}")

    result = re.sub(
        quoted_sheet_prefix_regex(capture_group="sheet1")
        + r"(?P<c1>[A-Z]{1,3})(?P<r1>\d+):"
        + quoted_sheet_prefix_regex(capture_group="sheet2")
        + r"(?P<c2>[A-Z]{1,3})(?P<r2>\d+)",
        replace_quoted,
        formula,
        flags=re.IGNORECASE,
    )

    def replace_unquoted(m: re.Match[str]) -> str:
        sheet1, sheet2 = m.group("sheet1"), m.group("sheet2")
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        if sheet1 != sheet2:
            return m.group(0)
        return format_range_key(sheet1, f"{c1}{int(r1)}", f"{c2}{int(r2)}")

    return re.sub(
        r"(?<![A-Za-z_'])(?P<sheet1>[A-Za-z][A-Za-z0-9_]*)!"
        r"(?P<c1>[A-Z]{1,3})(?P<r1>\d+):"
        r"(?P<sheet2>[A-Za-z][A-Za-z0-9_]*)!"
        r"(?P<c2>[A-Z]{1,3})(?P<r2>\d+)",
        replace_unquoted,
        result,
        flags=re.IGNORECASE,
    )


def normalize_excel_formula_with_name_state(
    formula: str,
    current_sheet: str,
    *,
    replacements: dict[str, str],
    names_re: re.Pattern[str] | None,
) -> str:
    """Normalize *formula* using pre-built defined-name replacement state."""
    if not formula or not formula.startswith("="):
        return formula
    masked, literals = _mask_string_literals(formula)
    result = _normalize_excel_formula_base(masked, current_sheet)
    result = _apply_named_range_replacements(result, replacements, names_re)
    result = _collapse_same_sheet_range_prefixes(result)
    return _unmask_string_literals(result, literals)


def expand_defined_names(
    formula: str,
    *,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
    name_state: NamedRangeReplacementState | None = None,
) -> str:
    """Replace defined names with sheet-qualified A1, leaving `$` markers intact.

    Replacement targets are emitted with `$` on both axes so defined names stay
    absolute when parsed with `parse_preserving_axes`. String literals are
    masked so names inside quotes are not expanded.

    Pass `name_state` to reuse a compiled alternation regex instead of
    rebuilding it on every formula.
    """
    if not formula:
        return formula
    if name_state is None:
        name_state = build_named_range_replacement_state(named_ranges, named_range_ranges)
    masked, literals = _mask_string_literals(formula)
    result = _apply_named_range_replacements(
        masked, name_state.absolute_replacements, name_state.names_re
    )
    return _unmask_string_literals(result, literals)


def _abs_a1_coord(coord: str) -> str:
    from fastpyxl.utils.cell import coordinate_from_string

    col_letter, row = coordinate_from_string(coord.replace("$", ""))
    return f"${col_letter}${row}"


def _absolutize_defined_name_target(target: str) -> str:
    """Rewrite a defined-name A1 target so both axes are `$`-absolute."""
    sheet, rest = parse_address(target)
    prefix = quote_sheet_if_needed(sheet)
    if ":" in rest:
        start, end = rest.split(":", 1)
        if "!" in end:
            end_sheet, end_coord = parse_address(end)
            return f"{prefix}!{_abs_a1_coord(start)}:{quote_sheet_if_needed(end_sheet)}!{_abs_a1_coord(end_coord)}"
        return f"{prefix}!{_abs_a1_coord(start)}:{_abs_a1_coord(end)}"
    return f"{prefix}!{_abs_a1_coord(rest)}"


def normalize_excel_formula(
    formula: str,
    current_sheet: str,
    *,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> str:
    """Regex-normalize a formula string (`=...`) for fallback parsing.

    Transitional and **not** the `normalized_formula` dialect. Prefer AST
    `render_formula` (`A1_ABSOLUTE`) whenever a tree is available.

    - Same-sheet refs (`A1`) become `Sheet!A1` using *current_sheet*.
    - Same-sheet ranges become single-prefix (`Sheet!A1:A3`); cross-sheet
      ranges keep both endpoints sheet-qualified.
    - Resolves defined names when maps are provided.
    - Strips `$` markers and qualifies range endpoints.
    """
    if not formula or not formula.startswith("="):
        return formula
    state = build_named_range_replacement_state(named_ranges, named_range_ranges)
    return normalize_excel_formula_with_name_state(
        formula,
        current_sheet,
        replacements=state.replacements,
        names_re=state.names_re,
    )


@dataclass(frozen=True, slots=True)
class PreparedFormula:
    """Result of preparing a cell formula for AST parsing."""

    normalized_formula: str


def prepare_formula(
    formula: str,
    current_sheet: str,
    *,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> PreparedFormula:
    """Regex-normalize *formula* for the cell on *current_sheet*.

    Transitional helper. Prefer `parse_preserving_axes` plus `render_formula`
    when an AST is available.
    """
    return PreparedFormula(
        normalized_formula=normalize_excel_formula(
            formula,
            current_sheet,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
    )
