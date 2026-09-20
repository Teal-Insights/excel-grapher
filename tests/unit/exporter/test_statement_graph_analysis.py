"""Statement-graph labels, series drilldown, and NetworkX export."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.semantic_catalog import SemanticCatalogError, load_semantic_catalog
from excel_grapher.exporter.semantic_graph import build_statement_graph
from excel_grapher.exporter.semantic_viz import to_semantic_viz_payload, write_semantic_viz_html
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


def _labelled_a12_bindings() -> dict:
    document = _a12_bindings()
    document["series"][0]["notes"] = "Stock path with mixed member formulas."
    document["series"][0]["sdmx_notes"] = "Illustrative TIME_PERIOD series."
    document["series"][0]["groups"] = [{"path": ["Engine", "Stocks"]}]
    for concept in document["concept_scheme"]["concepts"]:
        if concept["id"] == "OBS_VALUE":
            concept["name"] = "Observation value"
            concept["description"] = "Measured path value"
    return document


def _labelled_zipper_bindings() -> dict:
    document = _zipper_bindings()
    by_id = {entry["id"]: entry for entry in document["series"]}
    by_id["debt"]["notes"] = "Debt stock zipper."
    by_id["adjustment"]["notes"] = "Interest adjustment zipper."
    return document


def _view(workbook: Path, document: dict):
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    bindings = validate_bindings_document(document)
    view = load_semantic_catalog(graph, bindings, workbook=workbook)
    return view, graph, catalog


def test_statement_nodes_copy_binding_labels_and_keep_one_shape_each(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    view, graph, catalog = _view(workbook, _labelled_a12_bindings())
    statement_graph = build_statement_graph(view, graph)
    series = catalog.get("path")
    assert len(series.statements) == 3
    nodes = [node for node in statement_graph.nodes if node.series_id == "path"]
    assert len(nodes) == 3
    assert len({node.shape_key for node in nodes}) == 3
    for node in nodes:
        assert node.labels.notes == "Stock path with mixed member formulas."
        assert node.labels.sdmx_notes == "Illustrative TIME_PERIOD series."
        assert node.labels.compute_name == "compute_path"
        assert node.labels.groups == (("Engine", "Stocks"),)
        assert node.labels.measure_concept == "OBS_VALUE"
        assert node.labels.measure_name == "Observation value"
        assert node.labels.measure_description == "Measured path value"
        assert node.labels.axis_labels is None


def test_payload_json_includes_nested_labels(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    document = _labelled_a12_bindings()
    view, graph, _catalog = _view(workbook, document)
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(document), workbook=workbook, view=view
    )
    data = payload.to_dict()
    nodes = [node for node in data["nodes"] if node["series_id"] == "path"]
    assert nodes
    for node in nodes:
        assert "formula" not in node
        labels = node["labels"]
        assert labels["notes"] == "Stock path with mixed member formulas."
        assert labels["compute_name"] == "compute_path"
        assert labels["groups"] == [["Engine", "Stocks"]]
        assert labels["measure_name"] == "Observation value"


def test_html_viewer_embeds_labels_and_shape_key(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    document = _labelled_a12_bindings()
    view, graph, _catalog = _view(workbook, document)
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(document), workbook=workbook, view=view
    )
    html_path = tmp_path / "path.html"
    write_semantic_viz_html(payload, html_path)
    html = html_path.read_text(encoding="utf-8")
    assert "Stock path with mixed member formulas." in html
    assert "shape_key" in html
    assert "compute_path" in html


def test_remainder_node_has_empty_labels(tmp_path: Path) -> None:
    from excel_grapher.grapher.node import make_cell_node

    workbook = _a12_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _labelled_a12_bindings())
    graph.add_node(make_cell_node("Engine", "Z", 99, value=1, is_leaf=True))
    statement_graph = build_statement_graph(view, graph)
    remainder = next(node for node in statement_graph.nodes if node.is_remainder)
    assert remainder.labels.notes is None
    assert remainder.labels.groups == ()
    assert remainder.labels.compute_name is None


def test_drilldown_series_returns_cells_formulas_and_adjacency(tmp_path: Path) -> None:
    from excel_grapher.exporter.semantic_drilldown import drilldown_series

    workbook = _a12_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _labelled_a12_bindings())
    drilldown = drilldown_series(view, graph, "path")
    assert drilldown.series_id == "path"
    assert drilldown.statement_id is None
    assert len(drilldown.statements) == 3
    assert [cell.address for cell in drilldown.cells] == ["Engine!A2", "Engine!B2", "Engine!C2"]
    formulas = [cell.formula for cell in drilldown.cells]
    assert all(formulas)
    assert len(set(formulas)) == 3
    assert all(cell.in_series for cell in drilldown.cells)
    assert all(cell.statement_id is not None for cell in drilldown.cells)
    assert {cell.statement_id for cell in drilldown.cells} == {"path__0", "path__1", "path__2"}
    consumers = {edge.consumer_cell for edge in drilldown.edges}
    assert "Engine!B2" in consumers
    assert "Engine!C2" in consumers
    payload = drilldown.to_dict()
    assert payload["series_id"] == "path"
    assert [cell["address"] for cell in payload["cells"]] == [
        "Engine!A2",
        "Engine!B2",
        "Engine!C2",
    ]


def test_drilldown_statement_is_a_subset_of_the_series(tmp_path: Path) -> None:
    from excel_grapher.exporter.semantic_drilldown import drilldown_series, drilldown_statement

    workbook = _a12_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _labelled_a12_bindings())
    series = drilldown_series(view, graph, "path")
    one = drilldown_statement(view, graph, "path__1")
    assert one.series_id == "path"
    assert one.statement_id == "path__1"
    assert [cell.address for cell in one.cells] == ["Engine!B2"]
    assert one.cells[0].formula == series.cells[1].formula
    assert all(
        edge.consumer_cell == "Engine!B2" or edge.producer_cell == "Engine!B2" for edge in one.edges
    )


def test_drilldown_unknown_series_fails_closed(tmp_path: Path) -> None:
    from excel_grapher.exporter.semantic_drilldown import drilldown_series, drilldown_statement

    workbook = _a12_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _a12_bindings())
    with pytest.raises(SemanticCatalogError, match="unknown series"):
        drilldown_series(view, graph, "missing")
    with pytest.raises(SemanticCatalogError, match="unknown statement"):
        drilldown_statement(view, graph, "missing")


def test_zipper_drilldown_keeps_cross_series_neighbors(tmp_path: Path) -> None:
    from excel_grapher.exporter.semantic_drilldown import drilldown_series

    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _labelled_zipper_bindings())
    debt = drilldown_series(view, graph, "debt")
    neighbor_ids = {cell.series_id for cell in debt.neighbors}
    assert "adjustment" in neighbor_ids
    assert all(not cell.in_series for cell in debt.neighbors)
    pairs = {(edge.consumer_id, edge.producer_id, edge.access) for edge in debt.edges}
    assert ("debt", "adjustment", "identity") in pairs
    assert any(consumer == "adjustment" and producer == "debt" for consumer, producer, _ in pairs)


def test_statement_graph_to_networkx_is_a_multidigraph(tmp_path: Path) -> None:
    pytest.importorskip("networkx")
    from excel_grapher.exporter.semantic_graph import statement_graph_to_networkx

    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _labelled_zipper_bindings())
    statement_graph = build_statement_graph(view, graph)
    nx_graph = statement_graph_to_networkx(statement_graph)
    assert nx_graph.is_directed()
    assert nx_graph.is_multigraph()
    series_ids = {
        data["series_id"] for _, data in nx_graph.nodes(data=True) if not data.get("is_remainder")
    }
    assert series_ids >= {"debt", "adjustment"}
    debt_nodes = [
        node_id for node_id, data in nx_graph.nodes(data=True) if data.get("series_id") == "debt"
    ]
    assert debt_nodes
    assert nx_graph.nodes[debt_nodes[0]]["notes"] == "Debt stock zipper."
    assert "formula" not in nx_graph.nodes[debt_nodes[0]]
    accesses = {
        data["access"]
        for _src, _dst, data in nx_graph.edges(data=True)
        if data.get("access") is not None
    }
    assert "identity" in accesses
    assert "shift" in accesses
    method_graph = statement_graph.to_networkx()
    assert method_graph.number_of_nodes() == nx_graph.number_of_nodes()
    assert method_graph.number_of_edges() == nx_graph.number_of_edges()


def test_drilldown_to_networkx_includes_formulas(tmp_path: Path) -> None:
    pytest.importorskip("networkx")
    from excel_grapher.exporter.semantic_drilldown import drilldown_series

    workbook = _a12_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _labelled_a12_bindings())
    nx_graph = drilldown_series(view, graph, "path").to_networkx()
    assert nx_graph.is_directed()
    assert "Engine!A2" in nx_graph
    assert "Engine!B2" in nx_graph
    assert nx_graph.nodes["Engine!B2"]["formula"]
    assert nx_graph.nodes["Engine!B2"]["in_series"] is True
    assert nx_graph.has_edge("Engine!B2", "Engine!A2")
