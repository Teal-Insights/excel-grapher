"""Tests for optional formula AST pre-parsing during graph extraction."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import fastpyxl
import pytest

from excel_grapher import DependencyGraph, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.formula_ast import (
    AstNode,
    FormulaParseError,
    bind_axes,
    parse,
    parse_preserving_axes,
)
from excel_grapher.grapher.node import make_cell_node
from excel_grapher.grapher.preparsed_formulas import warm_preparsed_formulas


def test_warm_preparsed_formulas_reuses_per_node_formula_ast(tmp_path: Path) -> None:
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

    def counting_parse(formula: str) -> AstNode:
        nonlocal parse_calls
        parse_calls += 1
        return parse(formula)

    with patch("excel_grapher.grapher.preparsed_formulas.parse", counting_parse):
        warmed = warm_preparsed_formulas(graph)

    assert parse_calls == 0
    assert len(warmed) == 1
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    formula = node.normalized_formula
    assert formula is not None
    assert formula in warmed
    assert node.formula_ast is not None
    assert node.address is not None
    assert warmed[formula] == bind_axes(node.formula_ast, node.address)


def test_warm_preparsed_formulas_parses_when_formula_ast_missing() -> None:
    sheet, coord = parse_address("S!B1")
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    graph = DependencyGraph()
    graph.add_node(
        Node(
            sheet=sheet,
            column=col,
            row=row,
            formula="=1+1",
            normalized_formula="=1+1",
            value=None,
            is_leaf=False,
        )
    )
    internal = graph._get_internal_node("S!B1")
    assert internal is not None
    internal.formula_ast = None
    internal._unparseable_formula = "=1+1"
    parse_calls = 0

    def counting_parse(formula: str) -> AstNode:
        nonlocal parse_calls
        parse_calls += 1
        return parse(formula)

    with patch("excel_grapher.grapher.preparsed_formulas.parse", counting_parse):
        warmed = warm_preparsed_formulas(graph)

    assert parse_calls == 1
    assert warmed["=1+1"] == parse("=1+1")


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
    default_node = graph_default.get_node("Sheet1!A2")
    assert default_node is not None
    assert default_node.formula_ast is not None

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


def test_warm_preparsed_formulas_raises_on_invalid_formula() -> None:
    sheet, coord = parse_address("S!A1")
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    graph = DependencyGraph()
    graph.add_node(
        Node(
            sheet=sheet,
            column=col,
            row=row,
            formula="=1+",
            normalized_formula="=1+",
            value=None,
            is_leaf=False,
        )
    )

    with pytest.raises(FormulaParseError):
        warm_preparsed_formulas(graph)


def test_warm_preparsed_formulas_binds_relative_ast_under_absolute_key() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "A", 1, value=10, is_leaf=True))
    rel = parse_preserving_axes("=A1*2", anchor="S!B1")
    abs_ast = parse_preserving_axes("=$A$1*2", anchor="S!C1")
    graph.add_node(make_cell_node("S", "B", 1, is_leaf=False, formula_ast=rel))
    graph.add_node(make_cell_node("S", "C", 1, is_leaf=False, formula_ast=abs_ast))

    b1 = graph.get_node("S!B1")
    c1 = graph.get_node("S!C1")
    assert b1 is not None and c1 is not None
    assert b1.normalized_formula == c1.normalized_formula
    nf = b1.normalized_formula
    assert nf is not None

    warmed = warm_preparsed_formulas(graph)
    assert warmed[nf] == bind_axes(rel, "S!B1")
    assert warmed[nf] == parse(nf)
    assert warmed[nf] == bind_axes(abs_ast, "S!C1")
