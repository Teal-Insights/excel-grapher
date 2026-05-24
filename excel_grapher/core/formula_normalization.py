"""Canonical Excel formula normalization for sheet-qualified A1 references.

Graph extraction, evaluation, and codegen share these rules so bare cell
references resolve against the formula cell's sheet and named ranges expand
consistently before :mod:`excel_grapher.core.formula_ast` parsing.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import cast

from excel_grapher.core.address_keys import format_cell_key, normalize_range_key

_FUNC_LIKE = frozenset({"IF", "OR", "AND", "NOT", "SUM", "MAX", "MIN", "AVG"})


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


def _normalize_excel_formula_base(formula: str, current_sheet: str) -> str:
    """Strip $ markers, qualify ranges and cells, without defined-name substitution."""
    result = formula

    def replace_quoted_range(m: re.Match[str]) -> str:
        return normalize_range_key(m.group(0), current_sheet=current_sheet)

    result = re.sub(
        r"'(?P<sheet>[^']+)'!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_quoted_range,
        result,
    )

    def replace_unquoted_range(m: re.Match[str]) -> str:
        return normalize_range_key(m.group(0), current_sheet=current_sheet)

    result = re.sub(
        r"(?<![A-Za-z_'])(?P<sheet>[A-Za-z][A-Za-z0-9_]*)!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)\s*:\s*\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)",
        replace_unquoted_range,
        result,
    )

    def replace_local_range(m: re.Match[str]) -> str:
        return normalize_range_key(m.group(0), current_sheet=current_sheet)

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
    result = _normalize_excel_formula_base(formula, current_sheet)
    return _apply_named_range_replacements(result, replacements, names_re)


def normalize_excel_formula(
    formula: str,
    current_sheet: str,
    *,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> str:
    """Normalize a formula string (``=...``) for transpilation and parsing.

    - Same-sheet refs (``A1``) become ``Sheet!A1`` using *current_sheet*.
    - Resolves defined names when maps are provided.
    - Strips ``$`` markers and qualifies range endpoints.
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
