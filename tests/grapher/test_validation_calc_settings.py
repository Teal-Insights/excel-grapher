from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import xlsxwriter

from excel_grapher import get_calc_settings


def _patch_workbook_calcpr(
    src: Path,
    dst: Path,
    *,
    iterate: bool,
    iterate_count: int,
    iterate_delta: float,
) -> None:
    """
    Patch xl/workbook.xml calcPr attributes by rewriting the .xlsx zip.
    """
    with zipfile.ZipFile(src, "r") as zin:
        items = {name: zin.read(name) for name in zin.namelist()}

    root = ET.fromstring(items["xl/workbook.xml"])
    calc_pr = None
    for node in root.iter():
        if node.tag.endswith("calcPr"):
            calc_pr = node
            break
    if calc_pr is None:
        # Create <calcPr> under the workbook root (best-effort).
        calc_pr = ET.SubElement(root, "calcPr")

    calc_pr.attrib["iterate"] = "1" if iterate else "0"
    calc_pr.attrib["iterateCount"] = str(iterate_count)
    calc_pr.attrib["iterateDelta"] = str(iterate_delta)

    items["xl/workbook.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)


def test_get_calc_settings_reads_iterate_settings(tmp_path: Path) -> None:
    base = tmp_path / "base.xlsx"
    wb = xlsxwriter.Workbook(base)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    wb.close()

    patched = tmp_path / "patched.xlsx"
    _patch_workbook_calcpr(base, patched, iterate=True, iterate_count=42, iterate_delta=0.001)

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

