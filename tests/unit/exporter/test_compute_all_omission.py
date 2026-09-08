"""Omit public `compute_all` when `include_compute_all=False` (#457, #764)."""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.exporter.codegen import CodeGenerator


def _graph() -> DependencyGraph:
    graph = DependencyGraph()
    graph.add_node(
        Node(sheet="S", column="A", row=1, formula=None, value=1.0, is_leaf=True),
    )
    graph.add_node(
        Node(
            sheet="S",
            column="B",
            row=1,
            formula="=S!A1+1",
            normalized_formula="=S!A1+1",
            value=None,
            is_leaf=False,
        ),
    )
    graph.add_edge("S!B1", "S!A1")
    return graph


def test_generate_emits_compute_all_by_default() -> None:
    source = CodeGenerator(_graph()).generate(["S!B1"])
    assert "def compute_all(" in source
    assert "TARGETS = {" in source


def test_generate_omits_compute_all_when_disabled() -> None:
    source = CodeGenerator(_graph()).generate(["S!B1"], include_compute_all=False)
    assert "def compute_all(" not in source
    assert "TARGETS = {" not in source
    assert "def make_context(" in source
