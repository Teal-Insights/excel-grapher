"""ABS: evaluator matches live Excel on small synthetic workbooks (integration, slow).

Complements `test_abs_parity.py` (evaluator ↔ codegen). Requires xlwings or WSL/COM;
skips cleanly when Excel automation is unavailable.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.utils.excel_live_parity import LiveExcelCell, assert_evaluator_matches_live_excel


@pytest.mark.slow
def test_abs_excel_parity_literal_nested_and_cell_ref(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", value=-4),
            LiveExcelCell("B1", formula="=ABS(-3)"),
            LiveExcelCell("B2", formula="=ABS(A1)"),
            LiveExcelCell("B3", formula="=SUM(ABS(A1),10)"),
        ),
        targets=("S!B1", "S!B2", "S!B3"),
        workbook_stem="abs_numeric",
    )


@pytest.mark.slow
def test_abs_excel_parity_error_propagation_and_value_error(tmp_path: Path) -> None:
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("A1", formula="=1/0"),
            LiveExcelCell("B1", formula='=ABS("not a number")'),
            LiveExcelCell("B2", formula="=ABS(A1)"),
        ),
        targets=("S!B1", "S!B2"),
        workbook_stem="abs_errors",
    )


@pytest.mark.slow
def test_abs_excel_parity_sigma_band_pattern(tmp_path: Path) -> None:
    """Mirror ``normdist_sigma_band`` IF/ABS nesting from the sandbox workbook."""
    assert_evaluator_matches_live_excel(
        tmp_path=tmp_path,
        sheet="S",
        cells=(
            LiveExcelCell("J1", value=0.5),
            LiveExcelCell(
                "M1",
                formula='=IF(ABS(J1)<=1,"Within 1σ",IF(ABS(J1)<=2,"Within 2σ","Outlier >2σ"))',
            ),
            LiveExcelCell("J2", value=2.5),
            LiveExcelCell(
                "M2",
                formula='=IF(ABS(J2)<=1,"Within 1σ",IF(ABS(J2)<=2,"Within 2σ","Outlier >2σ"))',
            ),
        ),
        targets=("S!M1", "S!M2"),
        workbook_stem="abs_sigma_band",
    )
