"""AST-first per-node formula storage (#542)."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import fastpyxl

import excel_grapher.evaluator.evaluator as evaluator_module
from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.formula_ast import BinaryOpNode, CellRefNode, NumberNode, parse
from excel_grapher.core.formula_shape import intern_formula_shapes
from excel_grapher.grapher.cache import (
    GRAPH_CACHE_SCHEMA_VERSION,
    dependency_graph_from_json,
    dependency_graph_to_json,
)
from excel_grapher.grapher.formula_shapes import warm_formula_shapes
from excel_grapher.grapher.node import copy_node, make_cell_node


def _workbook_with_shared_formula(tmp_path: Path) -> Path:
    path = tmp_path / "ast_first.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["B1"].value = "=A1*2"
    ws["B2"].value = "=A1*2"
    ws["C1"].value = "=B1+1"
    wb.save(path)
    wb.close()
    return path


def test_extraction_leaves_formula_ast_none_when_unparseable(tmp_path: Path) -> None:
    path = tmp_path / "implicit_intersection.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["B1"].value = 2
    ws["C1"].value = "=SUM(IF(@A1:A3>0,@B1:B3,0))"
    wb.save(path)
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert node.normalized_formula is not None
    assert node.formula_ast is None
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!C1"], load_values=False)

    b1 = graph.get_node("Sheet1!B1")
    c1 = graph.get_node("Sheet1!C1")
    assert b1 is not None and c1 is not None
    assert b1.normalized_formula is not None
    assert c1.normalized_formula is not None
    assert b1.formula_ast == parse(b1.normalized_formula)
    assert c1.formula_ast == parse(c1.normalized_formula)
    assert graph.preparsed_formulas is None


def test_extraction_interns_identical_formula_asts(tmp_path: Path) -> None:
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!B2"], load_values=False)
    b1 = graph._get_internal_node("Sheet1!B1")
    b2 = graph._get_internal_node("Sheet1!B2")
    assert b1 is not None and b2 is not None
    assert b1.formula_ast is not None
    assert b1.formula_ast is b2.formula_ast


def test_copy_node_preserves_formula_ast() -> None:
    ast = parse("=Sheet1!A1+1")
    node = make_cell_node(
        "Sheet1",
        "B",
        1,
        formula="=A1+1",
        normalized_formula="=Sheet1!A1+1",
        is_leaf=False,
        formula_ast=ast,
    )
    cloned = copy_node(node)
    assert cloned.formula_ast is ast
    assert cloned.normalized_formula == "=Sheet1!A1+1"
    assert cloned is not node


def test_set_node_formula_parses_formula_ast() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "A", 1, is_leaf=True, value=1))
    graph.add_node(make_cell_node("Sheet1", "B", 1, is_leaf=False))
    graph.set_node_formula("Sheet1!B1", "=A1+2", "=Sheet1!A1+2")
    view = graph.get_node("Sheet1!B1")
    assert view is not None
    assert view.formula_ast == parse("=Sheet1!A1+2")

    graph.set_node_formula("Sheet1!B1", None, None)
    cleared = graph.get_node("Sheet1!B1")
    assert cleared is not None
    assert cleared.formula_ast is None


def test_json_cache_round_trips_formula_ast(tmp_path: Path) -> None:
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=True)
    assert GRAPH_CACHE_SCHEMA_VERSION >= 5

    restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    original = graph.get_node("Sheet1!C1")
    loaded = restored.get_node("Sheet1!C1")
    assert original is not None and loaded is not None
    assert loaded.normalized_formula == original.normalized_formula
    assert loaded.formula_ast == original.formula_ast
    assert loaded.formula_ast == parse(loaded.normalized_formula or "")


def test_evaluator_seeds_ast_cache_from_per_node_formula_ast(tmp_path: Path) -> None:
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B1"], load_values=True)
    assert graph.preparsed_formulas is None

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
        assert ev.evaluate(["Sheet1!B1"])["Sheet1!B1"] == 20.0
        assert parse_calls == 0
        assert ev.ast_cache_info().currsize >= 1


def test_formula_shape_bindings_are_keyed_by_node_key(tmp_path: Path) -> None:
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(
        path,
        ["Sheet1!B1", "Sheet1!B2", "Sheet1!C1"],
        load_values=False,
        warm_formula_shapes=True,
    )
    table = graph.formula_shapes
    assert table is not None
    assert "Sheet1!B1" in table.bindings
    assert "Sheet1!B2" in table.bindings
    assert table.lookup("Sheet1!B1") is not None
    assert table.lookup("=Sheet1!A1*2") is None

    b1 = table.lookup("Sheet1!B1")
    b2 = table.lookup("Sheet1!B2")
    assert b1 is not None and b2 is not None
    assert b1[0] == b2[0]
    assert b1[1] is b2[1]


def test_intern_formula_shapes_binds_each_node_key() -> None:
    table = intern_formula_shapes(
        [
            ("Sheet1!A1", "=Sheet1!B1+Sheet1!C1"),
            ("Sheet1!A2", "=Sheet1!B2+Sheet1!C2"),
            ("Sheet1!A3", "=Sheet1!B1+Sheet1!C1"),
            ("Sheet1!A4", "=SUM(Sheet1!A1:A3)"),
        ]
    )
    assert len(table.bindings) == 4
    assert len(table.shapes) == 2
    plus_a = table.lookup("Sheet1!A1")
    plus_b = table.lookup("Sheet1!A2")
    plus_dup = table.lookup("Sheet1!A3")
    assert plus_a is not None and plus_b is not None and plus_dup is not None
    assert plus_a[0] == plus_b[0] == plus_dup[0]
    assert plus_a[1] is plus_b[1]
    assert plus_a[2] == plus_dup[2]
    assert plus_a[2] != plus_b[2]
    assert table.lookup("=missing") is None


def test_warm_formula_shapes_uses_per_node_ast() -> None:
    graph_node = make_cell_node(
        "S",
        "B",
        1,
        normalized_formula="=S!A1+1",
        is_leaf=False,
        formula_ast=BinaryOpNode("+", CellRefNode("S!A1"), NumberNode(1.0)),
    )
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "A", 1, value=1, is_leaf=True))
    graph.add_node(graph_node)
    table = warm_formula_shapes(graph)
    found = table.lookup("S!B1")
    assert found is not None
    assert found[2] == (CellRefNode("S!A1"),)


def test_pickle_preserves_per_node_formula_ast(tmp_path: Path) -> None:
    import pickle

    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    restored = pickle.loads(pickle.dumps(graph))
    original = graph.get_node("Sheet1!C1")
    loaded = restored.get_node("Sheet1!C1")
    assert original is not None and loaded is not None
    assert loaded.formula_ast == original.formula_ast
    assert restored.formula_shapes is None
    assert restored.preparsed_formulas is None


def test_projection_copy_keeps_formula_ast(tmp_path: Path) -> None:
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    projected = graph._copy_for_projection()
    original = graph._get_internal_node("Sheet1!C1")
    cloned = projected._get_internal_node("Sheet1!C1")
    assert original is not None and cloned is not None
    assert cloned.formula_ast is original.formula_ast
