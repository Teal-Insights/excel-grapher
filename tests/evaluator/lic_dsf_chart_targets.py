"""
LIC-DSF Chart Data export targets (aligned with example/map_lic_dsf_indicators.py).

Used by slow integration tests to build the same dependency closure as the
indicator-mapping script: fixed signal ranges plus stress-test and figure rows.
"""

from __future__ import annotations

from pathlib import Path
from typing import Literal, TypedDict

import fastpyxl.utils.cell

from excel_grapher import format_cell_key
from excel_grapher.evaluator.name_utils import normalize_address

WORKBOOK_PATH = Path("example/data/lic-dsf-template-2025-08-12.xlsm")

# Figure 1 chart series: MX shock (U63) and Threshold (U66) — common parity probes.
_CHART_PARITY_SHORTLIST_RAW: list[str] = [
    "'Chart Data'!U63",
    "'Chart Data'!U66",
]


def chart_parity_shortlist_keys() -> list[str]:
    """Normalized keys for the small Chart Data parity slice (MX shock + Threshold column U)."""
    return [normalize_address(k) for k in _CHART_PARITY_SHORTLIST_RAW]


GRAPH_MAX_DEPTH = 50
GRAPH_USE_CACHED_DYNAMIC_REFS = True


class ExportRangeConfig(TypedDict):
    label: str
    range_spec: str
    entrypoint_mode: Literal["row_group", "per_cell"]


STRESS_TEST_ROW_LABELS: list[str] = [
    "Baseline",
    "A1. Key variables at their historical averages in 2024-2034 2/",
    "B1. Real GDP growth",
    "B2. Primary balance",
    "B3. Exports",
    "B4. Other flows 3/",
    "B5. Depreciation",
    "B6. Combination of B1-B5",
    "",
    "C1. Combined contingent liabilities",
    "C2. Natural disaster",
    "C3. Commodity price",
    "C4. Market Financing",
    "A2. Alternative Scenario :[Customize, enter title]",
]

STRESS_TEST_BLOCKS: list[tuple[str, int]] = [
    ("PV of Debt-to-GDP Ratio", 239),
    ("PV of Debt-to-Revenue Ratio", 281),
    ("Debt Service-to-Revenue Ratio", 318),
    ("Debt Service-to-GDP Ratio", 351),
]

FIGURE_DATA_ROWS: list[int] = [
    51,
    61,
    62,
    63,
    64,
    66,
    93,
    103,
    104,
    105,
    106,
    108,
    135,
    145,
    146,
    147,
    148,
    150,
    177,
    187,
    188,
    189,
    190,
    192,
    263,
    264,
    265,
    267,
    306,
    341,
    342,
    343,
]

EXPORT_FIXED_RANGES: list[ExportRangeConfig] = [
    {
        "label": "External DSA risk rating signals",
        "range_spec": "'Chart Data'!D10:D17",
        "entrypoint_mode": "per_cell",
    },
    {
        "label": "Fiscal (Total Public Debt) risk rating signals",
        "range_spec": "'Chart Data'!I10:I14",
        "entrypoint_mode": "per_cell",
    },
    {
        "label": "Applicable tailored stress test signals",
        "range_spec": "'Chart Data'!I17:I19",
        "entrypoint_mode": "row_group",
    },
    {
        "label": "Fiscal space for moderate risk category",
        "range_spec": "'Chart Data'!E25:E27",
        "entrypoint_mode": "row_group",
    },
    {
        "label": "Overall rating",
        "range_spec": "'Chart Data'!L10:L11",
        "entrypoint_mode": "row_group",
    },
]


def _export_chart_data_ranges() -> list[ExportRangeConfig]:
    out: list[ExportRangeConfig] = list(EXPORT_FIXED_RANGES)
    seen_row_specs = {entry["range_spec"] for entry in out}

    def add_chart_data_row(row: int, label: str) -> None:
        range_spec = f"'Chart Data'!D{row}:X{row}"
        if range_spec in seen_row_specs:
            return
        out.append(
            {
                "label": label,
                "range_spec": range_spec,
                "entrypoint_mode": "row_group",
            }
        )
        seen_row_specs.add(range_spec)

    for metric_label, start_row in STRESS_TEST_BLOCKS:
        for i, row_label in enumerate(STRESS_TEST_ROW_LABELS):
            if not row_label:
                continue
            row = start_row + i
            add_chart_data_row(row, f"{metric_label} - {row_label}")

    for row in FIGURE_DATA_ROWS:
        add_chart_data_row(row, f"Figure data row {row}")

    return out


EXPORT_RANGES: list[ExportRangeConfig] = _export_chart_data_ranges()


def parse_range_spec(spec: str) -> tuple[str, str]:
    if "!" not in spec:
        raise ValueError(f"Range spec must contain '!': {spec!r}")
    sheet_part, range_part = spec.split("!", 1)
    sheet_part = sheet_part.strip()
    if sheet_part.startswith("'") and sheet_part.endswith("'"):
        sheet_part = sheet_part[1:-1].replace("''", "'")
    return sheet_part, range_part.strip()


def cells_in_range(sheet: str, range_a1: str) -> list[str]:
    if ":" in range_a1:
        start_a1, end_a1 = range_a1.split(":", 1)
        start_a1 = start_a1.strip()
        end_a1 = end_a1.strip()
    else:
        start_a1 = end_a1 = range_a1.strip()

    c1, r1 = fastpyxl.utils.cell.coordinate_from_string(start_a1)
    c2, r2 = fastpyxl.utils.cell.coordinate_from_string(end_a1)
    start_col_idx = fastpyxl.utils.cell.column_index_from_string(c1)
    end_col_idx = fastpyxl.utils.cell.column_index_from_string(c2)
    rlo, rhi = (r1, r2) if r1 <= r2 else (r2, r1)
    clo, chi = (
        (start_col_idx, end_col_idx)
        if start_col_idx <= end_col_idx
        else (end_col_idx, start_col_idx)
    )

    out: list[str] = []
    for row in range(rlo, rhi + 1):
        for col_idx in range(clo, chi + 1):
            col_letter = fastpyxl.utils.cell.get_column_letter(col_idx)
            out.append(format_cell_key(sheet, col_letter, row))
    return out


def collect_chart_data_cell_keys() -> list[str]:
    """All sheet-qualified cell keys from EXPORT_RANGES (deduplicated, stable order)."""
    keys: list[str] = []
    seen: set[str] = set()
    for entry in EXPORT_RANGES:
        sheet_name, range_a1 = parse_range_spec(entry["range_spec"])
        for key in cells_in_range(sheet_name, range_a1):
            if key not in seen:
                seen.add(key)
                keys.append(key)
    return keys
