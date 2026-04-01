"""
LIC-DSF: GDP forecast bps shocks vs recalculated workbook (slow).

Applies multiplicative shocks ``v * (1 + bps/10000)`` to **Input 3** row **12**
columns **X:AR**, recalculates via ``modify_and_recalculate_workbook`` (xlwings /
PowerShell+COM / LibreOffice), then asserts ``FormulaEvaluator`` matches cached
values on selected **Chart Data** strips within ``rtol=1e-5``.
"""

from __future__ import annotations

import math
from pathlib import Path

import fastpyxl
import pytest
from fastpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_grapher import create_dependency_graph, format_cell_key
from excel_grapher.evaluator.name_utils import normalize_address
from tests.evaluator.excel_workbook_parity import assert_workbook_parity
from tests.evaluator.lic_dsf_chart_targets import WORKBOOK_PATH
from tests.utils.modify_and_recalculate import (
    ExcelRecalculationError,
    modify_and_recalculate_workbook,
)

pytestmark = pytest.mark.slow

RTOL = 1e-5
ATOL = 1e-9
GRAPH_MAX_DEPTH = 100

INPUT_SHEET = "Input 3 - Macro-Debt data(DMX)"
INPUT_ROW = 12
INPUT_COL_START = "X"
INPUT_COL_END = "AR"

CHART_SHEET = "Chart Data"
CHART_ROWS = (61, 62, 63, 103, 104, 105, 145, 146, 147, 187, 188, 189)
CHART_COL_START = "D"
CHART_COL_END = "X"

BPS_LEVELS = (-300, -200, -100, 100, 200, 300)


def _input_col_letters() -> list[str]:
    a = column_index_from_string(INPUT_COL_START)
    b = column_index_from_string(INPUT_COL_END)
    lo, hi = (a, b) if a <= b else (b, a)
    return [get_column_letter(i) for i in range(lo, hi + 1)]


def _collect_gdp_bps_modifications(workbook: Path, bps_signed: float) -> dict[str, float]:
    """Multiplicative shock: new_value = old * (1 + bps_signed/10000)."""
    factor_frac = bps_signed / 10_000.0
    wb = fastpyxl.load_workbook(str(workbook), data_only=True, read_only=True, keep_vba=True)
    try:
        if INPUT_SHEET not in wb.sheetnames:
            return {}
        ws = wb[INPUT_SHEET]
        q = INPUT_SHEET.replace("'", "''")
        mods: dict[str, float] = {}
        for col in _input_col_letters():
            addr = f"{col}{INPUT_ROW}"
            val = ws[addr].value
            if val is None or isinstance(val, bool):
                continue
            if not isinstance(val, (int, float)) or not math.isfinite(float(val)):
                continue
            new_v = float(val) * (1.0 + factor_frac)
            mods[f"'{q}'!${col}${INPUT_ROW}"] = new_v
        return mods
    finally:
        wb.close()


def _chart_strip_target_keys() -> list[str]:
    keys: list[str] = []
    c_lo = column_index_from_string(CHART_COL_START)
    c_hi = column_index_from_string(CHART_COL_END)
    clo, chi = (c_lo, c_hi) if c_lo <= c_hi else (c_hi, c_lo)
    for row in CHART_ROWS:
        for ci in range(clo, chi + 1):
            letter = get_column_letter(ci)
            keys.append(normalize_address(format_cell_key(CHART_SHEET, letter, row)))
    return keys


@pytest.fixture(scope="module")
def chart_strip_targets() -> list[str]:
    return _chart_strip_target_keys()


@pytest.mark.parametrize("bps", BPS_LEVELS, ids=lambda b: f"{b:+d}bps")
def test_gdp_bps_shock_chart_evaluator_matches_recalc(
    bps: int,
    chart_strip_targets: list[str],
    tmp_path: Path,
) -> None:
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")

    mods = _collect_gdp_bps_modifications(WORKBOOK_PATH, float(bps))
    if not mods:
        pytest.skip(f"No numeric cells to shock in {INPUT_SHEET!r} {INPUT_COL_START}{INPUT_ROW}:{INPUT_COL_END}{INPUT_ROW}")

    out_path = tmp_path / f"lic_dsf_gdp_shock_{bps:+d}bps.xlsm"
    try:
        modify_and_recalculate_workbook(WORKBOOK_PATH, out_path, mods)
    except (ExcelRecalculationError, RuntimeError, ImportError) as e:
        pytest.skip(f"Workbook recalculation not available: {e}")

    graph = create_dependency_graph(
        out_path,
        chart_strip_targets,
        load_values=True,
        max_depth=GRAPH_MAX_DEPTH,
        use_cached_dynamic_refs=True,
    )

    missing = [k for k in chart_strip_targets if graph.get_node(k) is None]
    assert not missing, f"Graph missing {len(missing)} target(s), e.g. {missing[:5]}"

    assert_workbook_parity(
        graph,
        chart_strip_targets,
        rtol=RTOL,
        atol=ATOL,
        fail_fast=True,
    )
