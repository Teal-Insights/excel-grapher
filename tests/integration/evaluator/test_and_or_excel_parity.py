"""AND/OR: evaluator matches live Excel on range scans (integration, slow).

Complements evaluator ↔ codegen AND/OR tests. Requires xlwings or WSL/COM;
skips cleanly when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_and_or_excel_parity_over_ranges_with_blanks_and_errors(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=True),
            LiveExcelCell("A2", formula="=1/0"),
            LiveExcelCell("A3", value=False),
            LiveExcelCell("B1", formula="=AND(A1:A3)"),
            LiveExcelCell("B2", formula="=OR(A1:A3)"),
            LiveExcelCell("C1", value=None),
            LiveExcelCell("C2", value=True),
            LiveExcelCell("C3", value=None),
            LiveExcelCell("D1", formula="=AND(C1:C3)"),
            LiveExcelCell("D2", formula="=OR(C1:C3)"),
        ),
        targets=("S!B1", "S!B2", "S!D1", "S!D2"),
        workbook_stem="and_or_ranges",
    )
