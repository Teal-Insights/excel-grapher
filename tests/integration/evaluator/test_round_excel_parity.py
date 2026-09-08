"""ROUND: evaluator matches live Excel on small synthetic workbooks (integration, slow).

Complements `test_round_parity.py` (evaluator ↔ codegen). Requires xlwings or WSL/COM;
skips cleanly when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_round_excel_parity_half_away_from_zero(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=1.25),
            LiveExcelCell("B1", formula="=ROUND(1.25,1)"),
            LiveExcelCell("B2", formula="=ROUND(-1.25,1)"),
            LiveExcelCell("B3", formula="=ROUND(2.5,0)"),
            LiveExcelCell("B4", formula="=ROUND(125,-1)"),
            LiveExcelCell("B5", formula="=ROUND(A1,1)"),
        ),
        targets=("S!B1", "S!B2", "S!B3", "S!B4", "S!B5"),
        workbook_stem="round_ties",
    )


@pytest.mark.slow
def test_round_excel_parity_error_propagation_and_value_error(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", formula="=1/0"),
            LiveExcelCell("B1", formula='=ROUND("not a number",0)'),
            LiveExcelCell("B2", formula="=ROUND(A1,0)"),
        ),
        targets=("S!B1", "S!B2"),
        workbook_stem="round_errors",
    )
