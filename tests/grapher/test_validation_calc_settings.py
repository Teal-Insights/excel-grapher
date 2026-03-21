from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import get_calc_settings
from tests.utils.workbook_xml import patch_workbook_calcpr


def test_get_calc_settings_reads_iterate_settings(tmp_path: Path) -> None:
    base = tmp_path / "base.xlsx"
    wb = xlsxwriter.Workbook(base)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    wb.close()

    patched = tmp_path / "patched.xlsx"
    patch_workbook_calcpr(base, patched, iterate=True, iterate_count=42, iterate_delta=0.001)

    settings = get_calc_settings(patched)
    assert settings.iterate_enabled is True
    assert settings.iterate_count == 42
    assert settings.iterate_delta == 0.001


def test_get_calc_settings_defaults_when_missing(tmp_path: Path) -> None:
    path = tmp_path / "plain.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    wb.close()

    settings = get_calc_settings(path)
    assert settings.iterate_enabled in {True, False}
    assert isinstance(settings.iterate_count, int)
    assert isinstance(settings.iterate_delta, float)

