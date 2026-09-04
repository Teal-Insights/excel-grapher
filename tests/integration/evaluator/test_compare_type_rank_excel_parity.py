"""Excel type-rank comparison: evaluator matches live Excel (#651).

Comparison never coerces across types. Number < text < logical. A blank cell
compares as `0`; the empty string is text. Requires xlwings or WSL/COM;
skips when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_compare_type_rank_excel_parity(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=True),
            LiveExcelCell("B1", value=100),
            LiveExcelCell("C1", formula="=A1>B1"),
            LiveExcelCell("A2", value="10"),
            LiveExcelCell("B2", value=10),
            LiveExcelCell("C2", formula="=A2=B2"),
            LiveExcelCell("A3", value=""),
            LiveExcelCell("B3", value=0),
            LiveExcelCell("C3", formula="=A3=B3"),
            LiveExcelCell("A4", value="abc"),
            LiveExcelCell("B4", value="ABC"),
            LiveExcelCell("C4", formula="=A4=B4"),
            LiveExcelCell("A5", value="a"),
            LiveExcelCell("B5", value=1),
            LiveExcelCell("C5", formula="=A5<B5"),
            LiveExcelCell("A6", value="10"),
            LiveExcelCell("B6", value=2),
            LiveExcelCell("C6", formula="=A6+B6"),
            LiveExcelCell("A7", value=True),
            LiveExcelCell("B7", value=1),
            LiveExcelCell("C7", formula="=A7=B7"),
        ),
        targets=("S!C1", "S!C2", "S!C3", "S!C4", "S!C5", "S!C6", "S!C7"),
        workbook_stem="compare_type_rank",
    )
