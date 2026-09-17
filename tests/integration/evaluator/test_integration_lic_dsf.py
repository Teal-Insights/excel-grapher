"""LIC-DSF workbook: evaluator matches Excel cached values on indicator strips (integration).

Loads the real `.xlsm` template, discovers formula cells, and compares
`FormulaEvaluator` output to last-saved Excel results. Chart Data slices and export
targets live in `test_lic_dsf_chart_parity.py` and `lic_dsf_chart_targets.py`.
"""

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import DependencyGraph, create_dependency_graph
from tests.utils.discover_formula_cells import discover_formula_cells_in_rows
from tests.utils.excel_workbook_parity import assert_workbook_parity

WORKBOOK_PATH = Path("tests/fixtures/lic_dsf/lic-dsf-template-2025-08-12.xlsm")

INDICATOR_CONFIG = {
    "B1_GDP_ext": [35, 36, 39, 40],
    "B3_Exports_ext": [35, 36, 39, 40],
    "B4_other flows_ext": [35, 36, 39, 40],
}

RTOL = 1e-5
ATOL = 1e-9


def _indicator_formula_cells() -> list[str]:
    """Return sheet-qualified formula addresses on the configured indicator rows."""
    targets: list[str] = []
    wb_f = fastpyxl.load_workbook(
        WORKBOOK_PATH,
        data_only=False,
        read_only=True,
        keep_vba=True,
        keep_formula_cache=True,
    )
    try:
        for sheet_name, rows in INDICATOR_CONFIG.items():
            targets.extend(
                discover_formula_cells_in_rows(
                    WORKBOOK_PATH,
                    sheet_name,
                    rows,
                    wb_formulas=wb_f,
                )
            )
    finally:
        wb_f.close()
    return targets


@pytest.fixture(scope="module")
def lic_dsf_indicator_graph() -> tuple[DependencyGraph, list[str]]:
    """Load the LIC-DSF workbook graph for indicator-strip formula cells."""
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")

    targets = _indicator_formula_cells()
    graph = create_dependency_graph(
        WORKBOOK_PATH,
        targets,
        load_values=True,
        max_depth=100,
        use_cached_dynamic_refs=True,
    )
    return graph, targets


@pytest.mark.slow
def test_lic_dsf_indicator_strips_evaluator_matches_excel_cache(
    lic_dsf_indicator_graph: tuple[DependencyGraph, list[str]],
) -> None:
    """Indicator-strip formulas match last-saved Excel cached values."""
    graph, targets = lic_dsf_indicator_graph
    assert targets
    assert_workbook_parity(graph, targets, rtol=RTOL, atol=ATOL, fail_fast=False)
