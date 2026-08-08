"""INDEX whole-axis (row/col = 0): evaluator matches live Excel (integration, slow).

Complements `test_index_parity.py` (evaluator ↔ codegen). Requires xlwings or WSL/COM;
skips cleanly when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_index_excel_parity_row_zero_whole_column(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=5),
            LiveExcelCell("A2", value=0),
            LiveExcelCell("A3", value=7),
            LiveExcelCell("B1", formula="=SUM(INDEX(A1:A3,0))"),
            LiveExcelCell("B2", formula="=MATCH(7,INDEX(A1:A3,0),0)"),
        ),
        targets=("S!B1", "S!B2"),
        workbook_stem="index_row_zero",
    )


@pytest.mark.slow
def test_index_excel_parity_zero_axis_on_2d(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=1),
            LiveExcelCell("A2", value=4),
            LiveExcelCell("A3", value=7),
            LiveExcelCell("B1", value=2),
            LiveExcelCell("B2", value=5),
            LiveExcelCell("B3", value=8),
            LiveExcelCell("C1", value=3),
            LiveExcelCell("C2", value=6),
            LiveExcelCell("C3", value=9),
            LiveExcelCell("E1", formula="=SUM(INDEX(A1:C3,0,2))"),
            LiveExcelCell("E2", formula="=SUM(INDEX(A1:C3,2,0))"),
            LiveExcelCell("E3", formula="=SUM(INDEX(A1:C3,0,0))"),
        ),
        targets=("S!E1", "S!E2", "S!E3"),
        workbook_stem="index_zero_2d",
    )


@pytest.mark.slow
def test_index_excel_parity_match_true_computed_array_idiom(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=0),
            LiveExcelCell("A2", value=0),
            LiveExcelCell("A3", value=7),
            LiveExcelCell("B1", formula="=MATCH(TRUE,INDEX((A1:A3<>0),0),0)"),
        ),
        targets=("S!B1",),
        workbook_stem="index_zero_match_true",
    )
