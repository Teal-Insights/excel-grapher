"""Hashable formula AST intern (#550)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.core.formula_ast import (
    AstNode,
    FunctionCallNode,
    NumberNode,
    intern_formula_ast,
    parse,
)
from excel_grapher.grapher.cache import (
    GRAPH_CACHE_SCHEMA_VERSION,
    dependency_graph_from_json,
    dependency_graph_to_json,
)


def test_function_call_args_are_frozen_tuples() -> None:
    args = [NumberNode(1.0)]
    node = FunctionCallNode("ABS", args)
    args.append(NumberNode(2.0))
    assert node.args == (NumberNode(1.0),)
    parsed = parse("=SUM(Sheet1!A1,1)")
    assert isinstance(parsed, FunctionCallNode)
    assert isinstance(parsed.args, tuple)


def test_formula_trees_are_hashable_and_intern_by_identity() -> None:
    first = parse("=SUM(Sheet1!A1,1)")
    second = parse("=SUM(Sheet1!A1,1)")
    assert first == second
    assert first is not second
    intern: dict[AstNode, AstNode] = {}
    interned_first = intern_formula_ast(first, intern)
    interned_second = intern_formula_ast(second, intern)
    assert interned_first is first
    assert interned_second is first
    assert interned_second is not second
    assert hash(first) == hash(second)


def test_extraction_interns_without_json_intern_keys(tmp_path: Path) -> None:
    path = tmp_path / "autofill.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["A2"].value = 20
    ws["B1"].value = "=A1*2"
    ws["B2"].value = "=A2*2"
    wb.save(path)
    wb.close()

    import excel_grapher.grapher.builder as builder

    assert not hasattr(builder, "ast_to_json")
    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!B2"], load_values=False)
    b1 = graph._get_internal_node("Sheet1!B1")
    b2 = graph._get_internal_node("Sheet1!B2")
    assert b1 is not None and b2 is not None
    assert b1.formula_ast is not None
    assert b1.formula_ast is b2.formula_ast


def test_json_cache_assigns_integer_formula_ast_ids(tmp_path: Path) -> None:
    path = tmp_path / "offset.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["A2"].value = 20
    ws["B1"].value = "=A1*2"
    ws["B2"].value = "=A1*2"
    wb.save(path)
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!B2"], load_values=False)
    assert GRAPH_CACHE_SCHEMA_VERSION >= 8
    payload = dependency_graph_to_json(graph)
    pool = payload["formula_asts"]
    assert isinstance(pool, list)
    assert len(pool) == 2
    ids = [
        node_payload["formula_ast_id"]
        for node_payload in payload["nodes"]
        if "formula_ast_id" in node_payload
    ]
    assert ids == [0, 1] or ids == [1, 0]
    assert len(set(ids)) == 2
    for node_payload in payload["nodes"]:
        assert "formula_ast" not in node_payload
        assert "formula_ast_key" not in node_payload

    restored = dependency_graph_from_json(payload)
    loaded_b1 = restored._get_internal_node("Sheet1!B1")
    loaded_b2 = restored._get_internal_node("Sheet1!B2")
    assert loaded_b1 is not None and loaded_b2 is not None
    assert loaded_b1.formula_ast is not loaded_b2.formula_ast
    original_b1 = graph.get_node("Sheet1!B1")
    assert original_b1 is not None
    assert loaded_b1.formula_ast == original_b1.formula_ast


def test_json_cache_rejects_non_list_formula_asts() -> None:
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.node import make_cell_node

    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "A", 1, is_leaf=True, value=1))
    payload = dependency_graph_to_json(graph)
    payload["formula_asts"] = {}
    with pytest.raises(TypeError, match="formula_asts"):
        dependency_graph_from_json(payload)


def test_json_cache_rejects_out_of_range_formula_ast_id(tmp_path: Path) -> None:
    path = tmp_path / "one.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["B1"].value = "=A1+1"
    wb.save(path)
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!B1"], load_values=False)
    payload = dependency_graph_to_json(graph)
    for node_payload in payload["nodes"]:
        if "formula_ast_id" in node_payload:
            node_payload["formula_ast_id"] = 99
    with pytest.raises(TypeError, match="formula_ast_id"):
        dependency_graph_from_json(payload)
