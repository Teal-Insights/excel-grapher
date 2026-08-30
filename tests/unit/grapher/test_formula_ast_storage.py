"""AST-first per-node formula storage (#542)."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import fastpyxl
import pytest

import excel_grapher.evaluator.evaluator as evaluator_module
from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    NumberNode,
    parse,
    parse_preserving_axes,
)
from excel_grapher.grapher.cache import (
    GRAPH_CACHE_SCHEMA_VERSION,
    dependency_graph_from_json,
    dependency_graph_to_json,
)
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.formula_shapes import warm_formula_shapes
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node, copy_node, make_cell_node


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


def test_extraction_stores_formula_ast(tmp_path: Path) -> None:
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!C1"], load_values=False)

    b1 = graph.get_node("Sheet1!B1")
    c1 = graph.get_node("Sheet1!C1")
    assert b1 is not None and c1 is not None
    assert b1.normalized_formula is not None
    assert c1.normalized_formula is not None
    assert b1.formula_ast == parse_preserving_axes("=A1*2", anchor="Sheet1!B1")
    assert c1.formula_ast == parse_preserving_axes("=B1+1", anchor="Sheet1!C1")
    assert graph.preparsed_formulas is None


def test_extraction_interns_identical_formula_asts(tmp_path: Path) -> None:
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!B2"], load_values=False)
    b1 = graph._get_internal_node("Sheet1!B1")
    b2 = graph._get_internal_node("Sheet1!B2")
    assert b1 is not None and b2 is not None
    assert b1.formula_ast is not None
    # Same raw `=A1*2` at B1 vs B2 is a different relative offset, so ASTs differ.
    assert b1.formula_ast != b2.formula_ast


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
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "A", 1, is_leaf=True, value=1))
    graph.add_node(make_cell_node("Sheet1", "B", 1, is_leaf=False))
    graph.set_node_formula("Sheet1!B1", "=A1+2", "=Sheet1!A1+2")
    view = graph.get_node("Sheet1!B1")
    assert view is not None
    assert view.formula_ast == parse_preserving_axes("=A1+2", anchor="Sheet1!B1")
    assert view.normalized_formula == "=Sheet1!A1+2"

    graph.set_node_formula("Sheet1!B1", None, None)
    cleared = graph.get_node("Sheet1!B1")
    assert cleared is not None
    assert cleared.formula_ast is None


def test_set_node_formula_leaves_formula_ast_unset_when_unparseable() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "B", 1, is_leaf=False))
    graph.set_node_formula(
        "Sheet1!B1",
        "=SUM(IF(@A1:A3>0,@B1:B3,0))",
        "=SUM(IF(@Sheet1!A1:A3>0,@Sheet1!B1:B3,0))",
    )
    view = graph.get_node("Sheet1!B1")
    assert view is not None
    assert view.normalized_formula == "=SUM(IF(@Sheet1!A1:A3>0,@Sheet1!B1:B3,0))"
    assert view.formula_ast is None


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
    assert loaded.formula_ast == parse_preserving_axes("=B1+1", anchor="Sheet1!C1")


def test_json_cache_interns_formula_asts_by_canonical_ast(tmp_path: Path) -> None:
    path = _workbook_with_shared_formula(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!B2"], load_values=False)
    b1 = graph.get_node("Sheet1!B1")
    b2 = graph.get_node("Sheet1!B2")
    assert b1 is not None and b2 is not None
    shared = b1.normalized_formula
    assert shared is not None and shared == b2.normalized_formula

    payload = dependency_graph_to_json(graph)
    pool = payload["formula_asts"]
    assert isinstance(pool, dict)
    keys = [
        node_payload.get("formula_ast_key")
        for node_payload in payload["nodes"]
        if node_payload.get("normalized_formula")
    ]
    assert all(isinstance(key, str) and key in pool for key in keys)
    assert len(set(keys)) == 2
    for node_payload in payload["nodes"]:
        assert "formula_ast" not in node_payload

    restored = dependency_graph_from_json(payload)
    loaded_b1 = restored._get_internal_node("Sheet1!B1")
    loaded_b2 = restored._get_internal_node("Sheet1!B2")
    assert loaded_b1 is not None and loaded_b2 is not None
    assert loaded_b1.formula_ast is not None
    assert loaded_b1.formula_ast != loaded_b2.formula_ast
    assert loaded_b1.formula_ast == parse_preserving_axes("=A1*2", anchor="Sheet1!B1")
    assert loaded_b2.formula_ast == parse_preserving_axes("=A1*2", anchor="Sheet1!B2")


def test_json_cache_rejects_non_object_formula_asts() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "A", 1, is_leaf=True, value=1))
    payload = dependency_graph_to_json(graph)
    payload["formula_asts"] = []
    with pytest.raises(TypeError, match="formula_asts"):
        dependency_graph_from_json(payload)


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


def _cell(key: str, formula: str | None, normalized: str | None, *, is_leaf: bool = False) -> Node:
    sheet, rest = key.split("!", 1)
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=normalized,
        value=None,
        is_leaf=is_leaf,
    )


def _direct_edge(graph: DependencyGraph, dependent: str, precedent: str) -> None:
    dep = graph.get_node(dependent)
    assert dep is not None
    normalized = dep.normalized_formula
    assert normalized is not None
    start = normalized.index(precedent)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=DependencyCause.direct_ref,
            direct_sites_normalized=((start, start + len(precedent)),),
        ),
    )


def test_identity_transit_rewrites_keep_formula_ast_aligned() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!C1", None, None, is_leaf=True))
    graph.add_node(_cell("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!C1"
    assert node.formula_ast == parse("=Sheet1!C1")


def test_optimal_inline_rewrites_keep_formula_ast_aligned() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!D1", None, None, is_leaf=True))
    graph.add_node(_cell("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1"))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" in removed
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.normalized_formula == "=(Sheet1!D1*2)+1"
    assert node.formula_ast == parse("=(Sheet1!D1*2)+1")


def test_identity_transit_leaves_formula_ast_unset_when_rewrite_unparseable() -> None:
    unparseable = "=SUM(IF(@Sheet1!A1:A3>0,@Sheet1!C1:C3,0))+Sheet1!B1"
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!C1", None, None, is_leaf=True))
    graph.add_node(_cell("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1"))
    graph.add_node(_cell("Sheet1!A1", unparseable, unparseable))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.normalized_formula == "=SUM(IF(@Sheet1!A1:A3>0,@Sheet1!C1:C3,0))+Sheet1!C1"
    assert node.formula_ast is None


def test_warm_formula_shapes_uses_per_node_ast() -> None:
    graph_node = make_cell_node(
        "S",
        "B",
        1,
        normalized_formula="=S!A1+1",
        is_leaf=False,
        formula_ast=BinaryOpNode("+", CellRefNode("S!A1"), NumberNode(1.0)),
    )

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
