"""Expand graph/codegen targets (cells, ranges, defined names) to concrete roots."""

from __future__ import annotations

from collections.abc import Iterable

import fastpyxl.utils.cell

from excel_grapher.core import address_keys as _address_keys

from .parser import expand_range, format_key


def split_range_target_on_colon(t: str) -> tuple[str, str] | None:
    """Split a sheet-qualified range target into (start_addr, end_addr).

    Handles colons embedded within quoted sheet names (`'It''s Data'!A1:B2`).
    Returns `None` if the target contains no top-level colon.
    """
    in_quote = False
    i = 0
    while i < len(t):
        ch = t[i]
        if ch == "'":
            if in_quote and i + 1 < len(t) and t[i + 1] == "'":
                i += 2
                continue
            in_quote = not in_quote
        elif ch == ":" and not in_quote:
            return t[:i], t[i + 1 :]
        i += 1
    return None


def expand_targets_to_roots(
    targets: Iterable[str],
    *,
    sheetnames: list[str],
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
    max_range_cells: int = 5000,
) -> list[tuple[str, str]]:
    """Expand mixed target inputs into concrete `(sheet, single_cell_a1)` roots.

    Accepted target forms:

    - Sheet-qualified single cells (`Sheet1!A1`, `'My Sheet'!A1`).
    - Sheet-qualified rectangular ranges (`Sheet1!A1:B2`,
      `Sheet1!A1:Sheet1!B2`, `'My Sheet'!A1:B2`). Expansion follows
      `expand_range` and `max_range_cells`.
    - Defined names (`MyCell`, `MyRange`) resolved against
      `named_ranges` (single cell) and `named_range_ranges` (rectangle).

    Returns roots in first-occurrence order with duplicates removed. Raises
    `ValueError` for unknown defined names, missing sheets, malformed
    sheet-qualified targets, and ranges that span multiple sheets.
    """
    seen: set[str] = set()
    roots: list[tuple[str, str]] = []

    def _emit(sheet: str, a1: str) -> None:
        key = format_key(sheet, a1)
        if key in seen:
            return
        seen.add(key)
        roots.append((sheet, a1))

    def _expand_rect(
        sheet: str,
        start_a1: str,
        end_a1: str,
        *,
        target_label: str,
    ) -> None:
        try:
            start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_a1)
            end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_a1)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"Invalid range coordinates in target {target_label!r}: {start_a1}:{end_a1}"
            ) from exc
        for dep_sheet, dep_a1 in expand_range(
            sheet=sheet,
            start_col=start_col,
            start_row=int(start_row),
            end_col=end_col,
            end_row=int(end_row),
            max_cells=max_range_cells,
        ):
            _emit(dep_sheet, dep_a1)

    def _require_sheet(sheet: str, *, target: str) -> None:
        if sheet not in sheetnames:
            raise ValueError(f"Sheet not found: {sheet}")

    for raw in targets:
        t = str(raw)
        if not t:
            raise ValueError("Target must be a non-empty string")

        if "!" in t:
            split = split_range_target_on_colon(t)
            if split is not None:
                start_addr, end_addr = split
                try:
                    sheet, start_a1 = _address_keys.parse_address(start_addr)
                except ValueError as exc:
                    raise ValueError(f"Invalid target address: {t}") from exc
                _require_sheet(sheet, target=t)
                if "!" in end_addr:
                    try:
                        end_sheet, end_a1 = _address_keys.parse_address(end_addr)
                    except ValueError as exc:
                        raise ValueError(f"Invalid target address: {t}") from exc
                    if end_sheet != sheet:
                        raise ValueError(f"Range target spans multiple sheets: {t}")
                else:
                    end_a1 = end_addr
                _expand_rect(sheet, start_a1, end_a1, target_label=t)
            else:
                try:
                    sheet, cell_part = _address_keys.parse_address(t)
                except ValueError as exc:
                    raise ValueError(f"Invalid target address: {t}") from exc
                _require_sheet(sheet, target=t)
                _emit(sheet, cell_part)
            continue

        cell_resolved = named_ranges.get(t)
        if cell_resolved is not None:
            sheet, a1 = cell_resolved
            if sheet not in sheetnames:
                raise ValueError(f"Sheet not found: {sheet} (resolved from defined name {t!r})")
            _emit(sheet, a1)
            continue

        range_resolved = named_range_ranges.get(t)
        if range_resolved is not None:
            sheet, start_a1, end_a1 = range_resolved
            if sheet not in sheetnames:
                raise ValueError(f"Sheet not found: {sheet} (resolved from defined name {t!r})")
            _expand_rect(sheet, start_a1, end_a1, target_label=t)
            continue

        raise ValueError(f"Target must be sheet-qualified or a known defined name: {t}")

    return roots
