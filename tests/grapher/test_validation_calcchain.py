from __future__ import annotations

import re
import zipfile
from pathlib import Path

import fastpyxl
import xlsxwriter

from excel_grapher import create_dependency_graph, validate_graph


def _make_simple_chain_xlsx(path: Path) -> None:
    """
    Use XlsxWriter for deterministic xlsx output.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 2)  # A1
    ws.write_number(1, 0, 3)  # A2
    ws.write_formula(2, 0, "=A1+A2", None, 5)  # A3 cached
    ws.write_formula(3, 0, "=A3*2", None, 10)  # A4 cached
    wb.close()


def _sheet_id_for_sheet1(xlsx_path: Path) -> str:
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8", errors="replace")
    # Find the sheetId for name="Sheet1"
    m = re.search(r'<sheet[^>]*name="Sheet1"[^>]*sheetId="(\d+)"', wb_xml)
    assert m is not None, "Could not find Sheet1 sheetId in workbook.xml"
    return m.group(1)


def _with_calcchain(src_xlsx: Path, dst_xlsx: Path, *, sheet_id: str, cell_refs: list[str]) -> None:
    """
    Copy xlsx zip entries and add xl/calcChain.xml with the provided formula cell refs.
    """
    calc = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ]
    for r in cell_refs:
        calc.append(f'  <c r="{r}" i="{sheet_id}"/>')
    calc.append("</calcChain>")
    calc_xml = "\n".join(calc).encode("utf-8")

    with zipfile.ZipFile(src_xlsx, "r") as zin, zipfile.ZipFile(dst_xlsx, "w") as zout:
        for item in zin.infolist():
            # Skip existing calcChain.xml if any
            if item.filename == "xl/calcChain.xml":
                continue
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("xl/calcChain.xml", calc_xml)


def test_validate_graph_gracefully_handles_missing_calcchain(tmp_path: Path) -> None:
    """
    Many generated workbooks won't have calcChain.xml. Validation should not crash.
    """
    excel_path = tmp_path / "no_calcchain.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 2
    ws["A2"].value = 3
    ws["A3"].value = "=A1+A2"
    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A3"], load_values=False)
    result = validate_graph(graph, excel_path)

    assert result.is_valid is True
    assert result.in_graph_not_in_chain == set()
    assert result.in_chain_not_in_graph == set()
    assert any("calcChain" in m for m in result.messages)


def test_validate_graph_compares_formula_cells_to_calcchain(tmp_path: Path) -> None:
    src = tmp_path / "src.xlsx"
    _make_simple_chain_xlsx(src)

    sheet_id = _sheet_id_for_sheet1(src)
    with_chain = tmp_path / "with_chain.xlsx"
    _with_calcchain(src, with_chain, sheet_id=sheet_id, cell_refs=["A3", "A4"])

    graph = create_dependency_graph(with_chain, ["Sheet1!A4"], load_values=False)
    result = validate_graph(graph, with_chain, scope={"Sheet1"})

    assert result.is_valid is True
    assert result.in_graph_not_in_chain == set()
    assert result.in_chain_not_in_graph == set()
