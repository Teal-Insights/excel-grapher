"""Tests for optional formula AST pre-parsing during graph extraction."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import fastpyxl

from excel_grapher import create_dependency_graph
from excel_grapher.core.formula_ast import AstNode
from excel_grapher.grapher.preparsed_formulas import warm_preparsed_formulas


def test_warm_preparsed_formulas_deduplicates_by_normalized_formula(tmp_path: Path) -> None:
    excel_path = tmp_path / "dup_formulas.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["B1"].value = "=A1*2"
    ws["B2"].value = "=A1*2"
    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!B1", "Sheet1!B2"], load_values=False)
    parse_calls = 0
    original_parse = warm_preparsed_formulas.__globals__["parse"]

    def counting_parse(formula: str) -> AstNode:
        nonlocal parse_calls
        parse_calls += 1
        return original_parse(formula)

    with patch("excel_grapher.grapher.preparsed_formulas.parse", counting_parse):
        warmed = warm_preparsed_formulas(graph)

    assert parse_calls == 1
    assert len(warmed) == 1
    nf = graph.get_node("Sheet1!B1")
    assert nf is not None
    assert nf.normalized_formula in warmed


def test_create_dependency_graph_warm_ast_cache_opt_in(tmp_path: Path) -> None:
    excel_path = tmp_path / "warm_flag.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["A2"].value = "=A1+1"
    wb.save(excel_path)
    wb.close()

    graph_default = create_dependency_graph(excel_path, ["Sheet1!A2"], load_values=False)
    assert graph_default.preparsed_formulas is None

    graph_warm = create_dependency_graph(
        excel_path,
        ["Sheet1!A2"],
        load_values=False,
        warm_ast_cache=True,
    )
    assert graph_warm.preparsed_formulas is not None
    node = graph_warm.get_node("Sheet1!A2")
    assert node is not None
    assert node.normalized_formula in graph_warm.preparsed_formulas
