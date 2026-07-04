"""EXP: evaluator matches live Excel on small synthetic workbooks (integration, slow).

Complements `test_exp_parity.py` (evaluator ↔ codegen). Requires xlwings or WSL/COM;
skips cleanly when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_exp_excel_parity_scalar_and_cell_ref(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=1.0),
            LiveExcelCell("B1", formula="=EXP(1)"),
            LiveExcelCell("B2", formula="=EXP(A1)"),
        ),
        targets=("S!B1", "S!B2"),
        workbook_stem="exp_numeric",
    )


@pytest.mark.slow
def test_exp_excel_parity_error_propagation_value_error_and_overflow(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", formula="=1/0"),
            LiveExcelCell("B1", formula='=EXP("not a number")'),
            LiveExcelCell("B2", formula="=EXP(A1)"),
            LiveExcelCell("A2", value=709.782),
            LiveExcelCell("B3", formula="=EXP(A2)"),
            LiveExcelCell("A3", value=710),
            LiveExcelCell("B4", formula="=EXP(A3)"),
        ),
        targets=("S!B1", "S!B2", "S!B3", "S!B4"),
        workbook_stem="exp_errors",
    )


@pytest.mark.slow
def test_exp_excel_parity_logistic_convergence_pattern(tmp_path: Path) -> None:
    """Mirror Q-CRAFT logistic convergence with a computed year column (issue #333 MCVE)."""
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=1.0),
            LiveExcelCell("B1", formula="=EXP(A1)"),
            LiveExcelCell("A2", value=0.5),
            LiveExcelCell("A3", value=15.0),
            LiveExcelCell("B2", formula="=1/(1+EXP(-A2*(B1-A3)))"),
        ),
        targets=("S!B1", "S!B2"),
        workbook_stem="exp_logistic",
    )
