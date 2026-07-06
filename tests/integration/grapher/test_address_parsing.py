"""Graph build and calcChain validation with apostrophe sheet names."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import create_dependency_graph, validate_graph


def _sheet_id_for_name(xlsx_path: Path, sheet_name: str) -> str:
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8", errors="replace")
    pattern = rf'<sheet[^>]*name="{re.escape(sheet_name)}"[^>]*sheetId="(\d+)"'
    match = re.search(pattern, wb_xml)
    assert match is not None, f"Could not find {sheet_name!r} sheetId in workbook.xml"
    return match.group(1)


def _with_calcchain(src_xlsx: Path, dst_xlsx: Path, *, sheet_id: str, cell_refs: list[str]) -> None:
    calc = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ]
    for cell_ref in cell_refs:
        calc.append(f'  <c r="{cell_ref}" i="{sheet_id}"/>')
    calc.append("</calcChain>")
    calc_xml = "\n".join(calc).encode("utf-8")

    with zipfile.ZipFile(src_xlsx, "r") as zin, zipfile.ZipFile(dst_xlsx, "w") as zout:
        for item in zin.infolist():
            if item.filename == "xl/calcChain.xml":
                continue
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("xl/calcChain.xml", calc_xml)


@pytest.mark.xfail(
    reason="Formula normalization corrupts quoted apostrophe sheet refs ('O''Neil' -> 'O'Neil')",
    strict=True,
)
def test_create_dependency_graph_handles_apostrophe_sheet_names(tmp_path: Path) -> None:
    sheet_name = "O'Neil"
    excel_path = tmp_path / "apostrophe_sheet.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet(sheet_name)
    ws.write_number(0, 0, 2)  # A1
    ws.write_formula(0, 1, "='O''Neil'!A1*2", None, 4)  # B1
    wb.close()

    graph = create_dependency_graph(excel_path, ["'O''Neil'!B1"], load_values=False)
    node = graph.get_node("'O''Neil'!B1")
    assert node is not None
    assert node.sheet == sheet_name
    assert "'O''Neil'!A1" in graph.get_dependencies("'O''Neil'!B1")


def test_validate_graph_scope_filters_apostrophe_sheet_names(tmp_path: Path) -> None:
    sheet_name = "O'Neil"
    src = tmp_path / "apostrophe_src.xlsx"
    wb = xlsxwriter.Workbook(src)
    ws = wb.add_worksheet(sheet_name)
    ws.write_number(0, 0, 1)  # A1
    ws.write_formula(0, 1, "='O''Neil'!A1+1", None, 2)  # B1
    wb.close()

    sheet_id = _sheet_id_for_name(src, sheet_name)
    with_chain = tmp_path / "apostrophe_chain.xlsx"
    _with_calcchain(src, with_chain, sheet_id=sheet_id, cell_refs=["B1"])

    graph = create_dependency_graph(with_chain, ["'O''Neil'!B1"], load_values=False)
    result = validate_graph(graph, with_chain, scope={sheet_name})

    assert result.is_valid is True
    assert result.in_graph_not_in_chain == set()
    assert result.in_chain_not_in_graph == set()
