"""Bindings-backed series graph: occupancy from the shared catalog."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.exporter import to_series_graph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import inverted_graph_parts
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _zipper_bindings,
    _zipper_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a12_formula_shape import (
    _a12_bindings,
    _a12_workbook,
)


def test_zipper_series_graph_is_a_two_cycle(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    catalog, _deps, graph = inverted_graph_parts(workbook, _zipper_bindings())
    bindings = validate_bindings_document(_zipper_bindings())

    series_graph = to_series_graph(graph, bindings, workbook=workbook, catalog=catalog)

    assert [node.series_id for node in series_graph.nodes] == ["debt", "adjustment"]
    edges = {(edge.source, edge.target): edge.edge_count for edge in series_graph.edges}
    assert set(edges) == {("debt", "adjustment"), ("adjustment", "debt")}
    assert edges[("debt", "adjustment")] >= 1
    assert edges[("adjustment", "debt")] >= 1
    assert ("debt", "debt") not in edges


def test_to_series_graph_builds_catalog_when_omitted(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    _catalog, _deps, graph = inverted_graph_parts(workbook, _zipper_bindings())
    bindings = validate_bindings_document(_zipper_bindings())

    series_graph = to_series_graph(graph, bindings, workbook=workbook)

    assert {node.series_id for node in series_graph.nodes} == {"debt", "adjustment"}


def test_formula_shape_series_stays_one_node(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    catalog, _deps, graph = inverted_graph_parts(workbook, _a12_bindings())
    bindings = validate_bindings_document(_a12_bindings())

    series_graph = to_series_graph(graph, bindings, workbook=workbook, catalog=catalog)
    path_nodes = [node for node in series_graph.nodes if node.series_id == "path"]
    assert len(path_nodes) == 1
    assert path_nodes[0].cell_count > 1
