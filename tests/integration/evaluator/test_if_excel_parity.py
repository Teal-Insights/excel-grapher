"""Array-style SUM(IF): evaluator matches live Excel (integration, slow).

Complements `test_if_parity.py` (evaluator ↔ codegen). Requires xlwings or
WSL/COM; skips cleanly when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_sum_if_array_excel_parity(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=-1),
            LiveExcelCell("A2", value=2),
            LiveExcelCell("A3", value=3),
            LiveExcelCell("B1", value=10),
            LiveExcelCell("B2", value=20),
            LiveExcelCell("B3", value=30),
            LiveExcelCell("C1", value=100),
            LiveExcelCell("C2", value=200),
            LiveExcelCell("C3", value=300),
            LiveExcelCell("E1", value=2),
            LiveExcelCell("D1", formula="=SUM(IF(A1:A3>0,A1:A3))"),
            LiveExcelCell("D2", formula="=SUM(IF(A1:A3>0,B1:B3,0))"),
            LiveExcelCell("D3", formula="=SUM(IF(A1:A3>0,B1:B3,C1:C3))"),
            LiveExcelCell("D4", formula="=SUM(IF(A1:A3=E1,B1:B3,0))"),
            LiveExcelCell("D5", formula="=SUM(IF(A1:A3>0,IF(B1:B3>15,C1:C3,0),0))"),
            LiveExcelCell("D6", formula="=SUMPRODUCT(IF(A1:A3>0,B1:B3,0))"),
        ),
        targets=("S!D1", "S!D2", "S!D3", "S!D4", "S!D5", "S!D6"),
        workbook_stem="sum_if_array",
    )
