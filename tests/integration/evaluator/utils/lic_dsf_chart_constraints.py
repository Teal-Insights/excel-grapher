"""Dynamic-ref constraints for the LIC-DSF Chart Data shortlist.

`CHART_SHORTLIST_CONSTRAINT_RANGES` lists leaf cells that feed OFFSET / INDEX /
INDIRECT arguments in the dependency closure of `chart_parity_shortlist_keys()`.
Each cell is frozen to its template value as a singleton `Literal` so
`create_dependency_graph(..., use_cached_dynamic_refs=False)` can resolve the
chart slice without combinatorial fallback.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import fastpyxl
from fastpyxl.worksheet.formula import ArrayFormula

from excel_grapher import DynamicRefConfig, DynamicRefError
from tests.integration.evaluator.utils.lic_dsf_chart_targets import (
    cells_in_range,
    parse_range_spec,
)

# Leaf rectangles discovered by iterative strict graph builds of
# `chart_parity_shortlist_keys()` (Chart Data U63 / U66). Grouped by sheet.
CHART_SHORTLIST_CONSTRAINT_RANGES: tuple[str, ...] = (
    # Country name selected on Input 1; MATCH/VLOOKUP into lookup and Trigger tables.
    "'Input 1 - Basics'!C7",
    # Language selector (English/French/Portuguese/Spanish).
    "START!K10",
    # Country lookup (name, HIPC, MDRI, code) and language lookup (BB:BC).
    "lookup!C4:G73",
    "lookup!BB4:BC7",
    # Market-financing trigger table: header row plus country names.
    "Trigger!AA4",
    "Trigger!AB4:AB74",
    "Trigger!AC4:AE4",
    # Classification matrix and country-info table referenced by INDEX/MATCH.
    "Classification!A16:DZ18",
    "Country_Information!A1:D80",
    # Padded remittance / GDP-USD tables consumed by CI INDEX rows.
    "data_GDPUSD!A1:Y7",
    "data_REM_FINAL!A1:AA7",
    # Commercial-terms block (COM) used by tailored-test INDEX.
    "COM!A3:A9",
    "COM!B2:B9",
    "COM!C3:G9",
    # Swap-curve labels and tenors on the blend-floating sheet.
    "'BLEND floating calculations WB'!I25:I39",
    "'BLEND floating calculations WB'!J10:J24",
    "'BLEND floating calculations WB'!K10:K39",
    "'BLEND floating calculations WB'!M11:M19",
    "'BLEND floating calculations WB'!M21",
    "'BLEND floating calculations WB'!M24",
    "'BLEND floating calculations WB'!M29",
    "'BLEND floating calculations WB'!M34",
    "'BLEND floating calculations WB'!M39",
    "'BLEND floating calculations WB'!N20",
    "'BLEND floating calculations WB'!N22:N23",
    "'BLEND floating calculations WB'!N25:N28",
    "'BLEND floating calculations WB'!N30:N33",
    "'BLEND floating calculations WB'!N35:N38",
    # Extra imported-data leaves outside the padded GDP/remittance tables.
    "'Imported data'!A105:B105",
    "'Imported data'!AG78",
    "'Imported data'!D92",
    "'Imported data'!D105",
    "'Imported data'!E91",
)


def expand_chart_shortlist_constraint_keys() -> list[str]:
    """Expand `CHART_SHORTLIST_CONSTRAINT_RANGES` to unique cell keys."""
    keys: list[str] = []
    seen: set[str] = set()
    for spec in CHART_SHORTLIST_CONSTRAINT_RANGES:
        sheet, a1 = parse_range_spec(spec)
        for key in cells_in_range(sheet, a1):
            if key in seen:
                continue
            seen.add(key)
            keys.append(key)
    return keys


def _raw_cell_value(raw: object) -> object:
    if isinstance(raw, ArrayFormula):
        return raw.text
    return raw


def _literal_for_value(value: object) -> Any:
    """Build a singleton `Literal` domain from a template leaf value."""
    if value is None or value == "":
        return Literal[None]
    if isinstance(value, str) and value.startswith("="):
        raise DynamicRefError(f"Dynamic-ref constraint leaf is a formula, not a leaf: {value!r}")
    if isinstance(value, (bool, int, float, str)):
        return Literal.__getitem__((value,))
    raise DynamicRefError(
        f"Unsupported constraint leaf value type {type(value).__name__}: {value!r}"
    )


def chart_shortlist_constraints(workbook_path: Path) -> dict[str, Any]:
    """Return a `Literal` schema for every chart-shortlist dynamic-ref leaf."""
    keep_vba = workbook_path.suffix.lower() == ".xlsm"
    wb = fastpyxl.load_workbook(workbook_path, data_only=False, keep_vba=keep_vba)
    try:
        schema: dict[str, Any] = {}
        for key in expand_chart_shortlist_constraint_keys():
            sheet, a1 = parse_range_spec(key)
            schema[key] = _literal_for_value(_raw_cell_value(wb[sheet][a1].value))
        return schema
    finally:
        wb.close()


def chart_shortlist_dynamic_refs(workbook_path: Path) -> DynamicRefConfig:
    """Build `DynamicRefConfig` covering `chart_parity_shortlist_keys()`."""
    return DynamicRefConfig.from_constraints(chart_shortlist_constraints(workbook_path), {})
