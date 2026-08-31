"""Tests for opt-in interned formula AST shapes on DependencyGraph."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import DependencyGraph, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.formula_ast import (
    FormulaParseError,
    parse_formula_text,
    parse_preserving_axes,
)
from excel_grapher.grapher.cache import dependency_graph_from_json, dependency_graph_to_json
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.formula_shapes import warm_formula_shapes
from excel_grapher.grapher.node import Node, make_cell_node


def _cell_node(address: str, formula: str | None = None, *, value: object = None) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    ast = parse_formula_text(formula, anchor=address) if formula else None
    return make_cell_node(
        sheet,
        col,
        int(row),
        formula=formula,
        value=value,
        is_leaf=formula is None,
        formula_ast=ast,
        normalized_formula=None if ast is not None else formula,
    )


def test_warm_formula_shapes_interns_shared_skeleton() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell_node("S!A1", value=1))
    graph.add_node(_cell_node("S!B1", "=A1+1"))
    graph.add_node(_cell_node("S!B2", "=A2+1"))
    table = warm_formula_shapes(graph)
    assert len(table.shapes) == 1
    assert len(table.bindings) == 2
    left = table.lookup("S!B1")
    right = table.lookup("S!B2")
    assert left is not None and right is not None
    assert left[0] == right[0]
    assert left[1] is right[1]
    assert left[2] == right[2]
    assert left[2] == (parse_preserving_axes("=A1", anchor="S!B1"),)


def test_warm_formula_shapes_interns_mixed_relative_and_absolute() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell_node("S!A1", value=1))
    graph.add_node(_cell_node("S!B1", "=A1+$A$1"))
    graph.add_node(_cell_node("S!B2", "=A2+$A$1"))
    table = warm_formula_shapes(graph)
    assert len(table.shapes) == 1
    left = table.lookup("S!B1")
    right = table.lookup("S!B2")
    assert left is not None and right is not None
    assert left[0] == right[0]
    assert left[1] is right[1]
    assert left[2] == right[2]
    assert left[2] == (
        parse_preserving_axes("=A1", anchor="S!B1"),
        parse_preserving_axes("=$A$1", anchor="S!B1"),
    )


def test_warm_formula_shapes_keeps_distinct_absolute_params() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell_node("S!A1", value=1))
    graph.add_node(_cell_node("S!A2", value=2))
    graph.add_node(_cell_node("S!B1", "=$A$1+1"))
    graph.add_node(_cell_node("S!B2", "=$A$2+1"))
    table = warm_formula_shapes(graph)
    assert len(table.shapes) == 1
    left = table.lookup("S!B1")
    right = table.lookup("S!B2")
    assert left is not None and right is not None
    assert left[0] == right[0]
    assert left[1] is right[1]
    assert left[2] != right[2]
    assert left[2] == (parse_preserving_axes("=$A$1", anchor="S!B1"),)
    assert right[2] == (parse_preserving_axes("=$A$2", anchor="S!B2"),)


def test_create_dependency_graph_warm_formula_shapes_opt_in(tmp_path: Path) -> None:
    excel_path = tmp_path / "shapes.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["B1"].value = "=A1+1"
    ws["B2"].value = "=A1+2"
    wb.save(excel_path)
    wb.close()

    default = create_dependency_graph(excel_path, ["Sheet1!B1", "Sheet1!B2"], load_values=False)
    assert default.formula_shapes is None
    assert default.preparsed_formulas is None
    b1 = default.get_node("Sheet1!B1")
    assert b1 is not None
    assert b1.formula_ast is not None

    shaped = create_dependency_graph(
        excel_path,
        ["Sheet1!B1", "Sheet1!B2"],
        load_values=False,
        warm_formula_shapes=True,
    )
    assert shaped.formula_shapes is not None
    assert shaped.preparsed_formulas is None
    assert len(shaped.formula_shapes.shapes) == 2

    both = create_dependency_graph(
        excel_path,
        ["Sheet1!B1", "Sheet1!B2"],
        load_values=False,
        warm_ast_cache=True,
        warm_formula_shapes=True,
    )
    assert both.preparsed_formulas is not None
    assert both.formula_shapes is not None


def test_set_node_formula_invalidates_shape_table() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell_node("S!A1", value=1))
    graph.add_node(_cell_node("S!B1", "=A1+1"))
    graph.formula_shapes = warm_formula_shapes(graph)
    graph.set_node_formula("S!B1", "=A1+2", "=S!A1+2")
    assert graph.formula_shapes is None


def test_warm_formula_shapes_raises_on_invalid_formula() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell_node("S!A1", "=1+"))
    with pytest.raises(FormulaParseError):
        warm_formula_shapes(graph)


def test_pickle_drops_formula_shapes() -> None:
    import pickle

    graph = DependencyGraph()
    graph.add_node(_cell_node("S!A1", value=1))
    graph.add_node(_cell_node("S!B1", "=A1+1"))
    graph.formula_shapes = warm_formula_shapes(graph)
    restored = pickle.loads(pickle.dumps(graph))
    assert restored.formula_shapes is None
    assert restored.get_node("S!B1") is not None


def test_json_cache_drops_formula_shapes() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell_node("S!A1", value=1))
    graph.add_node(_cell_node("S!B1", "=A1+1"))
    graph.formula_shapes = warm_formula_shapes(graph)
    restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    assert restored.formula_shapes is None
    assert restored.get_node("S!B1") is not None


def test_compress_identity_transits_invalidates_formula_shapes() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell_node("Sheet1!C1", value=42))
    graph.add_node(_cell_node("Sheet1!B1", "=C1"))
    graph.add_node(_cell_node("Sheet1!A1", "=B1"))
    direct = DependencyCause.direct_ref
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=direct))
    formula = "=Sheet1!B1"
    start = formula.index("Sheet1!B1")
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=direct,
            direct_sites_normalized=((start, start + len("Sheet1!B1")),),
        ),
    )
    graph.formula_shapes = warm_formula_shapes(graph)
    assert graph.formula_shapes is not None
    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    assert graph.formula_shapes is None
