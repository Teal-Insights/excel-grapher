"""Shared helpers for patching OOXML inside .xlsx test fixtures."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def patch_workbook_calcpr(
    src: Path,
    dst: Path,
    *,
    iterate: bool,
    iterate_count: int,
    iterate_delta: float,
) -> None:
    """Patch xl/workbook.xml calcPr attributes by rewriting the .xlsx zip."""
    with zipfile.ZipFile(src, "r") as zin:
        items = {name: zin.read(name) for name in zin.namelist()}

    root = ET.fromstring(items["xl/workbook.xml"])
    calc_pr = None
    for node in root.iter():
        if node.tag.endswith("calcPr"):
            calc_pr = node
            break
    if calc_pr is None:
        calc_pr = ET.SubElement(root, "calcPr")

    calc_pr.attrib["iterate"] = "1" if iterate else "0"
    calc_pr.attrib["iterateCount"] = str(iterate_count)
    calc_pr.attrib["iterateDelta"] = str(iterate_delta)

    items["xl/workbook.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
