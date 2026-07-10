"""Specialize a shared row-node formula template for a member column.

Varying slots are indices into `walk_template_cell_refs` — a deterministic
left-to-right walk of single-cell references with range spans masked so range
endpoints are never counted or rewritten.
"""

from __future__ import annotations

import re

from fastpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_grapher.grapher.range_compression.ref_parser import (
    AbsCellRef,
    parse_cell_refs_with_abs,
)

_A1_IN_SPAN_RE = re.compile(r"^(?P<acol>\$)?(?P<col>[A-Za-z]{1,3})(?P<arow>\$)?(?P<row>\d+)$")


def walk_template_cell_refs(normalized_template: str) -> list[AbsCellRef]:
    """Return single-cell refs in deterministic left-to-right order.

    Range references are excluded (their spans are masked before cell matching),
    so slot indices never point at range endpoints.

    Args:
        normalized_template: Formula text, typically starting with `=`.

    Returns:
        Cell refs with `$` markers and source spans preserved.
    """
    return parse_cell_refs_with_abs(normalized_template, default_sheet="")


def validate_varying_ref_slots(
    normalized_template: str,
    varying_ref_slots: tuple[int, ...],
) -> tuple[int, ...]:
    """Validate and canonicalize `varying_ref_slots` for a row template.

    Args:
        normalized_template: Shared formula for the row node.
        varying_ref_slots: Indices into `walk_template_cell_refs`.

    Returns:
        Deduplicated slot indices (first-seen order preserved).

    Raises:
        ValueError: If a slot is out of range or points at an absolute-column ref.
    """
    slots = tuple(dict.fromkeys(varying_ref_slots))
    refs = walk_template_cell_refs(normalized_template)
    for slot in slots:
        if slot < 0 or slot >= len(refs):
            raise ValueError(
                f"varying_ref_slots index {slot} out of range for template "
                f"with {len(refs)} cell ref(s)"
            )
        if refs[slot].is_absolute_col:
            raise ValueError(
                f"varying ref slot {slot} has absolute column "
                f"({normalized_template[refs[slot].span[0] : refs[slot].span[1]]!r}); "
                "column must be relative"
            )
    return slots


def specialize_template(
    normalized_template: str,
    *,
    varying_ref_slots: tuple[int, ...],
    column: str,
) -> str:
    """Rewrite column-varying cell refs in a row template for `column`.

    Only occurrences whose walk indices appear in `varying_ref_slots` are
    changed; static cell refs and all range text stay intact. Sheet qualifiers
    and row absolute/relative markers are preserved.

    Args:
        normalized_template: Shared formula for the row node.
        varying_ref_slots: Indices into `walk_template_cell_refs` that are
            column-parameterized.
        column: Member column letters (e.g. `E`).

    Returns:
        Specialized formula string for that member column.

    Raises:
        ValueError: If a slot index is out of range, a varying ref has an
            absolute column (`$D`), or `column` is not a valid Excel column.
    """
    if not varying_ref_slots:
        return normalized_template

    member_col = _normalize_column(column)
    refs = walk_template_cell_refs(normalized_template)
    unique_slots = validate_varying_ref_slots(normalized_template, varying_ref_slots)

    parts: list[str] = []
    cursor = len(normalized_template)
    for slot in sorted(unique_slots, reverse=True):
        ref = refs[slot]
        start, end = ref.span
        if end > cursor:
            raise ValueError(f"overlapping or invalid spans while specializing slot {slot}")
        parts.append(normalized_template[end:cursor])
        parts.append(_rewrite_cell_ref_span(normalized_template[start:end], member_col))
        cursor = start
    parts.append(normalized_template[:cursor])
    parts.reverse()
    return "".join(parts)


def _normalize_column(column: str) -> str:
    col = column.strip().upper()
    if col.startswith("$"):
        raise ValueError(f"member column must be relative letters, got {column!r}")
    try:
        index = column_index_from_string(col)
    except ValueError as exc:
        raise ValueError(f"invalid Excel column {column!r}") from exc
    return get_column_letter(index)


def _rewrite_cell_ref_span(span_text: str, member_col: str) -> str:
    """Replace only the A1 column letters inside a matched cell-ref span."""
    bang = span_text.rfind("!")
    prefix = span_text[: bang + 1] if bang >= 0 else ""
    a1 = span_text[bang + 1 :] if bang >= 0 else span_text
    match = _A1_IN_SPAN_RE.match(a1)
    if match is None:
        raise ValueError(f"cannot parse cell ref span {span_text!r}")
    if match.group("acol") is not None:
        raise ValueError(
            f"varying ref has absolute column ({span_text!r}); column must be relative"
        )
    row_abs = match.group("arow") is not None
    row = match.group("row")
    new_a1 = f"{member_col}${row}" if row_abs else f"{member_col}{row}"
    return f"{prefix}{new_a1}"
