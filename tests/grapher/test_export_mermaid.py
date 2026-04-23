from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import create_dependency_graph, to_mermaid


def _make_chain_xlsx(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 2
    ws["A2"].value = 3
    ws["A3"].value = "=A1+A2"
    ws["A4"].value = "=A3*2"
    wb.save(path)
    wb.close()


def _make_special_formula_xlsx(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 2
    ws["A2"].value = 3
    ws["A3"].value = "=SUM(A1,A2)"
    wb.save(path)
    wb.close()


def test_to_mermaid_contains_nodes_and_edges(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _make_chain_xlsx(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=False)
    mm = to_mermaid(graph, max_nodes=10)

    assert mm.startswith("flowchart TD")

    # Node ids should be sanitized (no !)
    assert "Sheet1_A4" in mm
    assert "Sheet1_A1" in mm

    # Edges are present
    assert "Sheet1_A4 --> Sheet1_A3" in mm
    assert "Sheet1_A3 --> Sheet1_A1" in mm
    assert "Sheet1_A3 --> Sheet1_A2" in mm

    # Formula nodes label with <br> line break before formula text
    assert "Sheet1!A4<br>=A3*2" in mm


def test_to_mermaid_can_omit_formula_labels(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _make_chain_xlsx(excel_path)
    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=False)
    mm = to_mermaid(graph, max_nodes=10, include_formula_on_nodes=False)
    assert "<br>" not in mm


def test_to_mermaid_truncates_formula(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _make_chain_xlsx(excel_path)
    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=False)
    mm = to_mermaid(graph, max_nodes=10, max_formula_length=4)
    assert "Sheet1!A4<br>=A3*..." in mm


def test_to_mermaid_quotes_label_with_special_formula_chars(tmp_path: Path) -> None:
    excel_path = tmp_path / "special_formula.xlsx"
    _make_special_formula_xlsx(excel_path)
    graph = create_dependency_graph(excel_path, ["Sheet1!A3"], load_values=False)
    mm = to_mermaid(graph, max_nodes=10)
    assert 'Sheet1_A3("Sheet1!A3<br>=SUM(A1,A2)")' in mm
