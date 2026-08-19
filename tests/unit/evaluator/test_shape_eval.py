"""Tests for shape-compiled FormulaEvaluator evaluation."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import fastpyxl

import excel_grapher.evaluator.evaluator as evaluator_module
from excel_grapher import DependencyGraph, FormulaEvaluator, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.types import XlError
from excel_grapher.grapher.formula_shapes import warm_formula_shapes


def _make_node(
    address: str,
    formula: str | None,
    value: object = None,
) -> Node:
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


def test_shape_eval_skips_parse_and_matches_string_path() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 10))
    graph.add_node(_make_node("S!A2", None, 20))
    graph.add_node(_make_node("S!B1", "=S!A1*2"))
    graph.add_node(_make_node("S!B2", "=S!A2*2"))
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")
    graph.formula_shapes = warm_formula_shapes(graph)

    parse_calls = 0
    original_parse = evaluator_module.parse

    def counting_parse(formula: str):
        nonlocal parse_calls
        parse_calls += 1
        return original_parse(formula)

    with patch.object(evaluator_module, "parse", counting_parse), FormulaEvaluator(graph) as ev:
        shaped = ev.evaluate(["S!B1", "S!B2"])
        assert parse_calls == 0
        assert shaped["S!B1"] == 20.0
        assert shaped["S!B2"] == 40.0
        assert len(ev._shape_fns) == 1

    graph.formula_shapes = None
    with FormulaEvaluator(graph) as baseline:
        plain = baseline.evaluate(["S!B1", "S!B2"])
    assert plain == shaped


def test_shape_eval_if_short_circuits_like_string_path() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 1))
    graph.add_node(_make_node("S!A2", None, 0))
    graph.add_node(_make_node("S!B1", None, 10))
    graph.add_node(_make_node("S!B2", None, 20))
    graph.add_node(_make_node("S!C1", "=IF(S!A1,S!B1,S!B2)"))
    graph.add_node(_make_node("S!C2", "=IF(S!A2,S!B1,S!B2)"))
    for parent, child in (
        ("S!C1", "S!A1"),
        ("S!C1", "S!B1"),
        ("S!C1", "S!B2"),
        ("S!C2", "S!A2"),
        ("S!C2", "S!B1"),
        ("S!C2", "S!B2"),
    ):
        graph.add_edge(parent, child)
    graph.formula_shapes = warm_formula_shapes(graph)

    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!C1", "S!C2"])
    assert result["S!C1"] == 10
    assert result["S!C2"] == 20


def test_shape_eval_if_does_not_evaluate_unused_error_branch() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 1))
    graph.add_node(_make_node("S!B1", None, 10))
    graph.add_node(_make_node("S!C1", "=1/0"))
    graph.add_node(_make_node("S!D1", "=IF(S!A1,S!B1,S!C1)"))
    graph.add_node(_make_node("S!A2", None, 0))
    graph.add_node(_make_node("S!D2", "=IF(S!A2,S!B1,S!C1)"))
    for parent, child in (
        ("S!D1", "S!A1"),
        ("S!D1", "S!B1"),
        ("S!D1", "S!C1"),
        ("S!D2", "S!A2"),
        ("S!D2", "S!B1"),
        ("S!D2", "S!C1"),
    ):
        graph.add_edge(parent, child)
    graph.formula_shapes = warm_formula_shapes(graph)

    with FormulaEvaluator(graph) as ev:
        taken = ev.evaluate("S!D1")
        assert taken == 10
        assert "S!C1" not in ev._cache
        unused = ev.evaluate("S!D2")
        assert unused == XlError.DIV


def test_create_graph_warm_formula_shapes_evaluator_parity(tmp_path: Path) -> None:
    path = tmp_path / "shape_eval.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 3
    ws["A2"].value = 5
    ws["B1"].value = "=A1+1"
    ws["B2"].value = "=A2+1"
    wb.save(path)
    wb.close()

    shaped = create_dependency_graph(
        path,
        ["Sheet1!B1", "Sheet1!B2"],
        load_values=True,
        warm_formula_shapes=True,
    )
    plain = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!B2"], load_values=True)
    with FormulaEvaluator(shaped) as ev_shape, FormulaEvaluator(plain) as ev_plain:
        assert ev_shape.evaluate(["Sheet1!B1", "Sheet1!B2"]) == ev_plain.evaluate(
            ["Sheet1!B1", "Sheet1!B2"]
        )
