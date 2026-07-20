"""Trailing-space text compare and MATCH: evaluator matches live Excel (slow).

Locks Excel behavior for GitHub #434: relational `=` / `<>` and exact
`MATCH(..., 0)` treat trailing ASCII spaces as significant. Requires xlwings
or WSL/COM; skips cleanly when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_compare_trailing_spaces_excel_parity_literals_and_cells(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value="High"),
            LiveExcelCell("A2", value="High "),
            LiveExcelCell("A3", value=" High"),
            LiveExcelCell("B1", formula='="High"="High "'),
            LiveExcelCell("B2", formula='="High"<>"High "'),
            LiveExcelCell("B3", formula='=" High"="High"'),
            LiveExcelCell("B4", formula='="high"="HIGH"'),
            LiveExcelCell("B5", formula='="high"="HIGH "'),
            LiveExcelCell("B6", formula="=A1=A2"),
            LiveExcelCell("B7", formula="=A1=A3"),
            LiveExcelCell("B8", formula="=A2=A2"),
        ),
        targets=(
            "S!B1",
            "S!B2",
            "S!B3",
            "S!B4",
            "S!B5",
            "S!B6",
            "S!B7",
            "S!B8",
        ),
        workbook_stem="compare_trailing_spaces",
    )


@pytest.mark.slow
def test_match_exact_trailing_spaces_excel_parity(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value="High"),
            LiveExcelCell("A2", value="High "),
            LiveExcelCell("B1", formula='=MATCH("High",A1:A1,0)'),
            LiveExcelCell("B2", formula='=MATCH("High",A2:A2,0)'),
            LiveExcelCell("B3", formula='=MATCH("High ",A1:A1,0)'),
            LiveExcelCell("B4", formula='=MATCH("High ",A2:A2,0)'),
            LiveExcelCell("B5", formula="=MATCH(A1,A2:A2,0)"),
        ),
        targets=("S!B1", "S!B2", "S!B3", "S!B4", "S!B5"),
        workbook_stem="match_trailing_spaces",
    )
