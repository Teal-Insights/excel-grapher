"""Patch workbook calculation settings in test fixture .xlsx files.

We have .xlsx files that we use as fixtures in our tests. In some cases, we want to patch
these with different calculation settings to see how those settings change test results.

An .xlsx is a ZIP of XML files. The module exposes patch_workbook_calcpr, which:

1. Reads the source workbook as a zip and loads xl/workbook.xml.
2. Finds or creates the calcPr element (Office Open XML “calculation properties”).
3. Sets iterative-calculation attributes: iterate (on/off), iterateCount, and iterateDelta.
4. Writes a new zip at dst with the updated xl/workbook.xml and everything else unchanged.

Note: The implementation still does a full read of the whole .xlsx: it loads every zip
member into a dict (zin.read(name) for each name), changes one XML blob, then writes a new
zip. A truly minimal-on-disk approach would stream-copy zip entries and only parse/replace
xl/workbook.xml; this helper trades that for simplicity.
"""

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
