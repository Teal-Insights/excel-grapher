from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph, to_graphviz, to_mermaid, to_networkx


def _make_if_guarded_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")

    ws.write_number(0, 2, 0)  # C1
    ws.write_formula(0, 0, "=IF($C$1=0,B1,1)", None, 1)  # A1
    ws.write_formula(0, 1, "=IF($C$1=1,A1,2)", None, 2)  # B1

    wb.close()


def test_graphviz_styles_guarded_edges_as_dashed(tmp_path: Path) -> None:
    excel_path = tmp_path / "if_guarded.xlsx"
    _make_if_guarded_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    dot = to_graphviz(graph)

    # Guarded edge should be dashed with a label containing the guard expression.
    assert '"Sheet1!A1" -> "Sheet1!B1"' in dot
    assert "style=dashed" in dot
    assert "label=" in dot


def test_mermaid_uses_dashed_arrow_for_guarded_edges(tmp_path: Path) -> None:
    excel_path = tmp_path / "if_guarded.xlsx"
    _make_if_guarded_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    mm = to_mermaid(graph, max_nodes=10)

    # Guarded edge should use dashed arrow syntax.
    assert "Sheet1_A1 -.-> Sheet1_B1" in mm or "Sheet1_A1 -. " in mm


def test_networkx_includes_guard_attr_on_guarded_edges(tmp_path: Path) -> None:
    excel_path = tmp_path / "if_guarded.xlsx"
    _make_if_guarded_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    G = to_networkx(graph)

    assert ("Sheet1!A1", "Sheet1!B1") in G.edges
    assert G.edges[("Sheet1!A1", "Sheet1!B1")].get("guard") is not None
