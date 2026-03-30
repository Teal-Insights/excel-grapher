from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph, to_networkx


def _make_chain_xlsx(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 2)  # A1
    ws.write_number(1, 0, 3)  # A2
    ws.write_formula(2, 0, "=A1+A2", None, 5)  # A3 cached
    ws.write_formula(3, 0, "=A3*2", None, 10)  # A4 cached
    wb.close()


def test_to_networkx_roundtrip_nodes_edges(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _make_chain_xlsx(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=True)
    G = to_networkx(graph)

    assert set(G.nodes) >= {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3", "Sheet1!A4"}
    assert ("Sheet1!A4", "Sheet1!A3") in G.edges
    assert ("Sheet1!A3", "Sheet1!A1") in G.edges
    assert ("Sheet1!A3", "Sheet1!A2") in G.edges

    # Sanity-check a couple attrs
    assert G.nodes["Sheet1!A4"]["is_leaf"] is False
    assert G.nodes["Sheet1!A1"]["is_leaf"] is True
    assert G.nodes["Sheet1!A4"]["value"] == 10
