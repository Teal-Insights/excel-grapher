"""Runtime cache eval scaffold: parity gate after ``xl_cell`` / ``xl_eval`` dedup.

Runs evaluator ↔ export checks across representative scenarios (formula workbook,
structural blanks, circular references) and asserts embedded export size stays
within the post-refactor line budget.
"""

from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import DependencyGraph, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator
from tests.integration.user_flows.formula_test import FORMULA_TARGETS, WORKBOOK_PATH
from tests.integration.utils.parity_harness import (
    CACHE_EVAL_SCAFFOLD_LINE_BUDGET,
    assert_cache_eval_scaffold_within_budget,
    assert_codegen_matches_evaluator,
    count_cache_eval_scaffold_lines,
)


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


def _make_graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_formula_test_cases_parity_gate() -> None:
    """formula_test_cases.xlsx G3:G9: evaluator ↔ export on nested IF and strings."""
    graph = create_dependency_graph(WORKBOOK_PATH, FORMULA_TARGETS, load_values=True)
    result = assert_codegen_matches_evaluator(graph, FORMULA_TARGETS)
    assert_cache_eval_scaffold_within_budget(result.generated_code)


def test_blank_range_index_parity_gate(tmp_path: Path) -> None:
    """Structural blank cells stay None inside INDEX ranges (issue #39)."""
    path = tmp_path / "blank_range_gate.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["B1"].value = 20
    ws["D1"].value = "=INDEX(A1:B3,1,1)"
    ws["E1"].value = "=INDEX(A1:B3,3,2)"
    wb.save(path)
    wb.close()

    blank = ("Sheet1!A2:B3",)
    graph = create_dependency_graph(
        path, ["Sheet1!D1", "Sheet1!E1"], load_values=True, blank_ranges=blank
    )
    result = assert_codegen_matches_evaluator(graph, ["Sheet1!D1", "Sheet1!E1"], blank_ranges=blank)
    assert result.evaluator_results["Sheet1!D1"] == 10
    assert result.evaluator_results["Sheet1!E1"] == 0
    assert_cache_eval_scaffold_within_budget(result.generated_code)


def test_circular_reference_parity_gate() -> None:
    """Direct self-cycle returns 0 with warning in evaluator and export."""
    graph = _make_graph(_make_node("S!A1", "=S!A1", None))
    result = assert_codegen_matches_evaluator(graph, ["S!A1"])
    assert result.evaluator_results["S!A1"] == 0
    assert result.generated_results["S!A1"] == 0
    assert_cache_eval_scaffold_within_budget(result.generated_code)


def test_cache_eval_scaffold_line_budget_on_minimal_export() -> None:
    """Standalone export embeds shared helper; xl_cell/xl_eval stay thin wrappers."""
    graph = _make_graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=S!A1+1", None),
    )
    code = CodeGenerator(graph).generate(["S!B1"])

    line_count = assert_cache_eval_scaffold_within_budget(code)
    assert line_count <= CACHE_EVAL_SCAFFOLD_LINE_BUDGET
    assert count_cache_eval_scaffold_lines(code) == line_count
    assert "_evaluate_address(ctx, address, obtain_fn" in code
    assert code.count("def _evaluate_address(") == 1
    assert "preserve_structural_blank=True" in code
    assert "preserve_structural_blank=False" in code
