"""Tests for seeding FormulaEvaluator AST cache from graph pre-parsing."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import fastpyxl

import excel_grapher.evaluator.evaluator as evaluator_module
from excel_grapher import FormulaEvaluator, create_dependency_graph


def test_evaluator_seeds_ast_cache_from_graph_preparsed_formulas(tmp_path: Path) -> None:
    excel_path = tmp_path / "seed_eval.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["B1"].value = "=A1*2"
    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!B1"],
        load_values=True,
        warm_ast_cache=True,
    )

    parse_calls = 0
    original_parse = evaluator_module.parse

    def counting_parse(formula: str):
        nonlocal parse_calls
        parse_calls += 1
        return original_parse(formula)

    with (
        FormulaEvaluator(graph) as ev,
        patch.object(evaluator_module, "parse", counting_parse),
    ):
        ev.evaluate(["Sheet1!B1"])
        assert parse_calls == 0
        assert ev.ast_cache_info().currsize >= 1
