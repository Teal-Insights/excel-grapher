"""Declared structural blank ranges interact with graph build and the evaluator (integration).

Graphs that honor declared blanks skip omitted rectangles, and INDEX over those
holes resolves as empty (issue #39).
"""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _index_over_blank_workbook(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["B1"].value = 20
    ws["D1"].value = "=INDEX(A1:B3,1,1)"
    ws["E1"].value = "=INDEX(A1:B3,3,2)"
    wb.save(path)
    wb.close()


def test_create_dependency_graph_skips_blank_range_nodes(tmp_path: Path) -> None:
    path = tmp_path / "blank_range.xlsx"
    _index_over_blank_workbook(path)

    blank = ("Sheet1!A2:B3",)
    graph = create_dependency_graph(
        path, ["Sheet1!D1", "Sheet1!E1"], load_values=True, blank_ranges=blank
    )

    for addr in ("Sheet1!A2", "Sheet1!B2", "Sheet1!A3", "Sheet1!B3"):
        assert addr not in graph

    assert "Sheet1!A1" in graph
    assert "Sheet1!B1" in graph
    assert "Sheet1!D1" in graph
    assert "Sheet1!E1" in graph

    with FormulaEvaluator(graph, blank_ranges=blank) as ev:
        results = ev.evaluate(["Sheet1!D1", "Sheet1!E1"])
    assert results["Sheet1!D1"] == 10
    assert results["Sheet1!E1"] == 0


def test_missing_cell_outside_declared_blank_still_keyerror() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", "=S!Z99", None))

    with FormulaEvaluator(graph) as ev, pytest.raises(KeyError):
        ev.evaluate("S!A1")

    with FormulaEvaluator(graph, blank_ranges=("S!Z99",)) as ev2:
        assert ev2.evaluate("S!A1") == 0
