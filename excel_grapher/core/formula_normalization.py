"""Canonical Excel formula normalization for sheet-qualified A1 references.

Graph extraction, evaluation, and codegen share these rules so bare cell
references resolve against the formula cell's sheet and named ranges expand
consistently before `excel_grapher.core.formula_ast` parsing.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import cast

from excel_grapher.core.address_keys import format_cell_key, quote_sheet_if_needed

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


def build_named_range_replacement_state(
    named_ranges: dict[str, tuple[str, str]] | None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None,
) -> tuple[dict[str, str], re.Pattern[str] | None]:
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
            start_ref = format_cell_key(sheet, m_start.group(1), int(m_start.group(2)))
            end_ref = format_cell_key(sheet, m_end.group(1), int(m_end.group(2)))
            replacements[name] = f"{start_ref}:{end_ref}"

    if not replacements:
        return {}, None

    names = cast(list[str], sorted(replacements, key=len, reverse=True))
    alt = "|".join(re.escape(n) for n in names)
    names_re = re.compile(rf"\b(?:{alt})\b(?!\s*!)")
    return replacements, names_re


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
        return _format_whole_column_ref(m.group("sheet"), m.group("col"))

    result = re.sub(
        r"'(?P<sheet>[^']+)'!\$?(?P<col>[A-Z]{1,3})\s*:\s*\$?(?P=col)\b",
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
        return _format_whole_row_ref(m.group("sheet"), int(m.group("row")))

    result = re.sub(
        r"'(?P<sheet>[^']+)'!\$?(?P<row>\d+)\s*:\s*\$?(?P=row)\b",
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

    def replace_quoted_range(m: re.Match[str]) -> str:
        sheet = m.group("sheet")
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        a = format_cell_key(sheet, c1, int(r1))
        b = format_cell_key(sheet, c2, int(r2))
        return f"{a}:{b}"

    result = re.sub(
        r"'(?P<sheet>[^']+)'!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_quoted_range,
        result,
    )

    def replace_unquoted_range(m: re.Match[str]) -> str:
        sheet = m.group("sheet")
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        return f"{sheet}!{c1}{r1}:{sheet}!{c2}{r2}"

    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_unquoted_range,
        result,
    )

    def replace_local_range(m: re.Match[str]) -> str:
        c1, r1, c2, r2 = m.group("c1"), m.group("r1"), m.group("c2"), m.group("r2")
        ref1 = format_cell_key(current_sheet, c1, int(r1))
        ref2 = format_cell_key(current_sheet, c2, int(r2))
        return f"{ref1}:{ref2}"

    result = re.sub(
        r"(?<![!A-Za-z0-9_])(?<!\$)\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)(?![A-Za-z0-9_])",
        replace_local_range,
        result,
    )

    def replace_quoted_cell(m: re.Match[str]) -> str:
        sheet, col, row = m.group("sheet"), m.group("col"), m.group("row")
        return format_cell_key(sheet, col, int(row))

    result = re.sub(
        r"'(?P<sheet>[^']+)'!\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)",
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
        r"(?<![!A-Za-z0-9_])(?<!\$)\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)(?![A-Za-z0-9_!'])",
        replace_local_cell,
        result,
    )

    return result


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
    return _unmask_string_literals(result, literals)


def normalize_excel_formula(
    formula: str,
    current_sheet: str,
    *,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> str:
    """Normalize a formula string (`=...`) for transpilation and parsing.

    - Same-sheet refs (`A1`) become `Sheet!A1` using *current_sheet*.
    - Resolves defined names when maps are provided.
    - Strips `$` markers and qualifies range endpoints.
    """
    if not formula or not formula.startswith("="):
        return formula
    repl, names_re = build_named_range_replacement_state(named_ranges, named_range_ranges)
    return normalize_excel_formula_with_name_state(
        formula, current_sheet, replacements=repl, names_re=names_re
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
    """Normalize *formula* for the cell on *current_sheet* and return a bundle."""
    return PreparedFormula(
        normalized_formula=normalize_excel_formula(
            formula,
            current_sheet,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
    )
