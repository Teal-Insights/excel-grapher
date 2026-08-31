"""AST `render_formula` is the `normalized_formula` / provenance-span dialect (#555).

Regex `FormulaNormalizer` must not supply character spans that are later sliced
out of the AST-rendered `Node.normalized_formula`.
"""

from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.formula_ast import BinaryOpNode, NumberNode
from excel_grapher.grapher.parser import FormulaNormalizer


def _write_formula_workbook(path: Path, formula: str, *, a1: object = 2, a2: object = 3) -> Path:
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = a1
    ws["A2"].value = a2
    ws["B1"].value = formula
    wb.save(path)
    wb.close()
    return path


def _assert_spans_slice_normalized(graph, dependent: str, precedent: str) -> None:
    node = graph.get_node(dependent)
    assert node is not None
    normalized = node.normalized_formula
    assert normalized is not None
    prov = graph.get_edge_attrs(dependent, precedent).provenance
    assert prov is not None
    assert prov.direct_sites_normalized
    for start, end in prov.direct_sites_normalized:
        assert normalized[start:end] == precedent


def test_paren_formula_spans_match_ast_render_not_regex(tmp_path: Path) -> None:
    """`=(A1)` renders as `=Sheet1!A1`; regex-normalized text still has the parens."""
    raw = "=(A1)"
    regex = FormulaNormalizer().normalize(raw, "Sheet1")
    assert regex == "=(Sheet1!A1)"

    graph = create_dependency_graph(
        _write_formula_workbook(tmp_path / "parens.xlsx", raw),
        ["Sheet1!B1"],
        capture_dependency_provenance=True,
    )
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!A1"
    assert node.normalized_formula != regex
    _assert_spans_slice_normalized(graph, "Sheet1!B1", "Sheet1!A1")
    assert FormulaEvaluator(graph).evaluate("Sheet1!B1") == 2


def test_float_literal_and_spaced_sum_spans_match_ast_render(tmp_path: Path) -> None:
    raw = "=SUM( A1 )+1.0"
    regex = FormulaNormalizer().normalize(raw, "Sheet1")
    graph = create_dependency_graph(
        _write_formula_workbook(tmp_path / "spaces.xlsx", raw),
        ["Sheet1!B1"],
        capture_dependency_provenance=True,
    )
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    assert node.normalized_formula == "=SUM(Sheet1!A1)+1"
    assert node.normalized_formula != regex
    _assert_spans_slice_normalized(graph, "Sheet1!B1", "Sheet1!A1")
    assert FormulaEvaluator(graph).evaluate("Sheet1!B1") == 3


def test_lowercase_bare_refs_are_extracted_via_ast(tmp_path: Path) -> None:
    graph = create_dependency_graph(
        _write_formula_workbook(tmp_path / "lower.xlsx", "=a1+a2"),
        ["Sheet1!B1"],
        capture_dependency_provenance=True,
    )
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!A1+Sheet1!A2"
    assert graph.get_dependencies("Sheet1!B1") == {"Sheet1!A1", "Sheet1!A2"}
    _assert_spans_slice_normalized(graph, "Sheet1!B1", "Sheet1!A1")
    _assert_spans_slice_normalized(graph, "Sheet1!B1", "Sheet1!A2")
    assert FormulaEvaluator(graph).evaluate("Sheet1!B1") == 5


def test_scientific_literal_parses_into_formula_ast(tmp_path: Path) -> None:
    graph = create_dependency_graph(
        _write_formula_workbook(tmp_path / "sci.xlsx", "=1e2+A1"),
        ["Sheet1!B1"],
        capture_dependency_provenance=True,
    )
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    ast = node.formula_ast
    assert isinstance(ast, BinaryOpNode)
    assert ast.left == NumberNode(100.0)
    assert node.normalized_formula == "=100+Sheet1!A1"
    _assert_spans_slice_normalized(graph, "Sheet1!B1", "Sheet1!A1")
    assert FormulaEvaluator(graph).evaluate("Sheet1!B1") == 102
