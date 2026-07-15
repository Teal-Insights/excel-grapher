"""COUNTIF: evaluator matches live Excel on criteria scan (integration, slow).

Complements evaluator ↔ codegen COUNTIF tests. Requires xlwings or WSL/COM;
skips cleanly when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_countif_excel_parity_skips_error_cells_in_range(tmp_path: Path) -> None:
    """Excel COUNTIF ignores error cells in the criteria range (does not propagate)."""
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=10),
            LiveExcelCell("A2", formula="=1/0"),
            LiveExcelCell("A3", value=20),
            LiveExcelCell("B1", formula='=COUNTIF(A1:A3,">5")'),
        ),
        targets=("S!B1",),
        workbook_stem="countif_skip_errors",
    )


@pytest.mark.slow
def test_countif_excel_parity_numeric_and_text_criteria(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=10),
            LiveExcelCell("A2", value="text"),
            LiveExcelCell("A3", value=20),
            LiveExcelCell("B1", formula='=COUNTIF(A1:A3,">5")'),
            LiveExcelCell("B2", formula='=COUNTIF(A1:A3,"text")'),
        ),
        targets=("S!B1", "S!B2"),
        workbook_stem="countif_criteria",
    )
