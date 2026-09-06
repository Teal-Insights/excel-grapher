"""INDEX/MATCH across mixed labels and blanks: evaluator matches live Excel."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_index_match_blank_padded_vectors_excel_parity(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="Engine",
        cells=(
            LiveExcelCell("A5", value="Note"),
            LiveExcelCell("A7", value="Title"),
            LiveExcelCell("A8", value="Bond"),
            LiveExcelCell("A9", value="Loan"),
            LiveExcelCell("A10", value="Equity"),
            LiveExcelCell("D5", value=2020),
            LiveExcelCell("E5", value=2021),
            LiveExcelCell("D8", value=10),
            LiveExcelCell("E8", value=11),
            LiveExcelCell("D9", value=20),
            LiveExcelCell("E9", value=21),
            LiveExcelCell("D10", value=30),
            LiveExcelCell("E10", value=31),
            LiveExcelCell("A228", value="Loan"),
            LiveExcelCell("B230", value=2020),
            LiveExcelCell("B231", value=2021),
            LiveExcelCell(
                "E230",
                formula="=INDEX($A$5:$E$10,MATCH(A$228,$A$5:$A$10,0),MATCH(B230,$A$5:$E$5,0))",
            ),
            LiveExcelCell(
                "E231",
                formula="=INDEX($A$5:$E$10,MATCH(A$228,$A$5:$A$10,0),MATCH(B231,$A$5:$E$5,0))",
            ),
            LiveExcelCell("A229", value="Missing"),
            LiveExcelCell(
                "E232",
                formula="=INDEX($A$5:$E$10,MATCH(A229,$A$5:$A$10,0),MATCH(B230,$A$5:$E$5,0))",
            ),
        ),
        targets=("Engine!E230", "Engine!E231", "Engine!E232"),
        workbook_stem="index_match_multi_owner",
    )
