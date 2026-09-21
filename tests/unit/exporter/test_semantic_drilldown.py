"""Partition and bundle drilldown plus statement-graph fixture coverage."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_grapher.exporter.semantic_catalog import SemanticCatalogError, load_semantic_catalog
from excel_grapher.exporter.semantic_graph import build_statement_graph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import inverted_graph_parts, write_workbook
from tests.unit.exporter.inverted_tree.test_affine_access import (
    _decimate_bindings,
    _decimate_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a3_take import (
    _punched_bindings,
    _punched_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _zipper_bindings,
    _zipper_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a12_formula_shape import (
    _a12_bindings,
    _a12_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a20_matrix_join import (
    _cross_country_legal_bindings,
    _cross_country_legal_sheets,
)
from tests.unit.exporter.inverted_tree.test_shape_a20_matrix_join import (
    _zipper_bindings as _a20_zipper_bindings,
)
from tests.unit.exporter.inverted_tree.test_shape_a20_matrix_join import (
    _zipper_workbook as _a20_zipper_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a22_guarded_residual import (
    _series_may_cycle_bindings,
    _series_may_cycle_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a22_shift_k import (
    _stride_k_bindings,
    _stride_k_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a27_range_aggregates import (
    range_sum_bindings,
    range_sum_workbook,
)


def _view(workbook: Path, document: dict):
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    bindings = validate_bindings_document(document)
    view = load_semantic_catalog(graph, bindings, workbook=workbook)
    return view, graph, catalog


def test_drilldown_partition_keeps_one_matrix_block(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_partition, drilldown_series

    view, graph, _catalog = _view(_a20_zipper_workbook(tmp_path), _a20_zipper_bindings())
    debt = drilldown_series(view, graph, "debt")
    france = drilldown_partition(view, graph, "debt", ("France",))
    kenya = drilldown_partition(view, graph, "debt", ("Kenya",))
    assert france.series_id == "debt"
    assert france.partition == ("France",)
    assert kenya.partition == ("Kenya",)
    assert {cell.address for cell in france.cells}.isdisjoint(
        {cell.address for cell in kenya.cells}
    )
    assert {cell.address for cell in france.cells} | {cell.address for cell in kenya.cells} == {
        cell.address for cell in debt.cells
    }
    assert all(cell.partition == ("France",) for cell in france.cells)
    assert all(edge.partition == ("France",) for edge in france.edges)
    json.dumps(france.to_dict())


def test_drilldown_partition_unknown_fails_closed(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_partition

    view, graph, _catalog = _view(_a20_zipper_workbook(tmp_path), _a20_zipper_bindings())
    with pytest.raises(SemanticCatalogError, match="unknown partition"):
        drilldown_partition(view, graph, "debt", ("Spain",))
    with pytest.raises(SemanticCatalogError, match="unknown series"):
        drilldown_partition(view, graph, "missing", ("France",))


def test_drilldown_bundle_lists_every_canonical_edge(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_bundle

    view, graph, _catalog = _view(_zipper_workbook(tmp_path), _zipper_bindings())
    statement_graph = build_statement_graph(view, graph)
    identity = next(
        bundle
        for bundle in statement_graph.bundles
        if bundle.access == "identity" and bundle.consumer_id.startswith("debt")
    )
    drilldown = drilldown_bundle(
        view,
        graph,
        consumer_id=identity.consumer_id,
        producer_id=identity.producer_id,
        access=identity.access,
        distance=identity.distance,
        coeff=identity.coeff,
        offset=identity.offset,
        guarded=identity.guarded,
    )
    assert drilldown.consumer_id == identity.consumer_id
    assert drilldown.producer_id == identity.producer_id
    assert len(drilldown.edges) == identity.instance_edge_count
    pairs = {(edge.consumer_cell, edge.producer_cell) for edge in drilldown.edges}
    assert (identity.representative_consumer_cell, identity.representative_producer_cell) in pairs
    assert pairs == set(identity.instance_edges)
    payload = drilldown.to_dict()
    json.dumps(payload)
    assert payload["instance_edge_count"] == identity.instance_edge_count
    assert len(payload["edges"]) == identity.instance_edge_count


def test_drilldown_bundle_unknown_fails_closed(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_bundle

    view, graph, _catalog = _view(_zipper_workbook(tmp_path), _zipper_bindings())
    with pytest.raises(SemanticCatalogError, match="unknown bundle"):
        drilldown_bundle(
            view,
            graph,
            consumer_id="missing",
            producer_id="debt",
            access="identity",
            distance=0,
        )


def test_formula_shape_breaks_are_distinct_statements(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_series, drilldown_statement

    view, graph, catalog = _view(_a12_workbook(tmp_path), _a12_bindings())
    path = catalog.get("path")
    assert len(path.statements) == 3
    series = drilldown_series(view, graph, "path")
    assert [stmt.statement_id for stmt in series.statements] == [
        stmt.statement_id for stmt in path.statements
    ]
    one = drilldown_statement(view, graph, path.statements[1].statement_id)
    assert [cell.address for cell in one.cells] == list(path.statements[1].cells)


def test_holes_are_absent_from_member_cells(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_series

    view, graph, catalog = _view(_punched_workbook(tmp_path), _punched_bindings())
    punched = catalog.get("output_punched")
    drilldown = drilldown_series(view, graph, "output_punched")
    periods = [dict(cell.key)["TIME_PERIOD"] for cell in drilldown.cells]
    assert periods == [1, 3]
    assert {cell.address for cell in drilldown.cells} == set(punched.cells)
    assert len(punched.cells) < 3


def test_identity_and_shift_bundles_on_zipper_scan(tmp_path: Path) -> None:
    view, graph, _catalog = _view(_zipper_workbook(tmp_path), _zipper_bindings())
    statement_graph = build_statement_graph(view, graph)
    accesses = {(bundle.access, bundle.distance > 0) for bundle in statement_graph.bundles}
    assert ("identity", False) in accesses
    assert any(access == "shift" for access, _lag in accesses)
    assert statement_graph.stats.bundle_count < statement_graph.stats.instance_edge_count


def test_stride_scan_keeps_self_lag_bundles(tmp_path: Path) -> None:
    view, graph, _catalog = _view(
        _stride_k_workbook(tmp_path, 8, 2, stem="viz_stride"), _stride_k_bindings(8)
    )
    statement_graph = build_statement_graph(view, graph)
    self_lags = [
        bundle
        for bundle in statement_graph.bundles
        if bundle.consumer_id.startswith("path")
        and bundle.producer_id.startswith("path")
        and bundle.access == "shift"
    ]
    assert self_lags
    assert {bundle.distance for bundle in self_lags} == {2}


def test_affine_decimate_is_one_typed_bundle(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_bundle

    view, graph, _catalog = _view(_decimate_workbook(tmp_path), _decimate_bindings())
    statement_graph = build_statement_graph(view, graph)
    affine = [
        bundle
        for bundle in statement_graph.bundles
        if bundle.access == "affine"
        and bundle.consumer_id.startswith("sampled")
        and bundle.producer_id.startswith("source")
    ]
    assert affine
    assert {bundle.coeff for bundle in affine} == {2}
    assert {bundle.offset for bundle in affine} == {0}
    pairs = {pair for bundle in affine for pair in bundle.instance_edges}
    assert pairs == {
        ("Engine!A3", "Engine!A2"),
        ("Engine!B3", "Engine!C2"),
        ("Engine!C3", "Engine!E2"),
    }
    first = affine[0]
    drilldown = drilldown_bundle(
        view,
        graph,
        consumer_id=first.consumer_id,
        producer_id=first.producer_id,
        access="affine",
        distance=first.distance,
        coeff=2,
        offset=0,
    )
    assert len(drilldown.edges) == first.instance_edge_count
    assert {(edge.consumer_cell, edge.producer_cell) for edge in drilldown.edges} <= pairs


def test_range_sum_uses_gather_or_whole_access(tmp_path: Path) -> None:
    view, graph, _catalog = _view(range_sum_workbook(tmp_path), range_sum_bindings())
    statement_graph = build_statement_graph(view, graph)
    accesses = {
        bundle.access
        for bundle in statement_graph.bundles
        if bundle.consumer_id.startswith("out") and bundle.producer_id.startswith("src")
    }
    assert accesses & {"gather", "whole"}


def test_cross_partition_reads_stay_typed(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_partition

    workbook = write_workbook(tmp_path / "a20_cross_ok.xlsx", _cross_country_legal_sheets())
    view, graph, _catalog = _view(workbook, _cross_country_legal_bindings())
    statement_graph = build_statement_graph(view, graph)
    cross = [bundle for bundle in statement_graph.bundles if bundle.access == "cross_partition"]
    assert cross
    kenya = drilldown_partition(view, graph, "path", ("Kenya",))
    assert any(edge.access == "cross_partition" for edge in kenya.edges)
    assert any(edge.producer_cell == "Engine!B2" for edge in kenya.edges)


def test_guarded_bundle_expands_to_guarded_instance_edges(tmp_path: Path) -> None:
    from excel_grapher.exporter import drilldown_bundle

    view, graph, _catalog = _view(
        _series_may_cycle_workbook(tmp_path), _series_may_cycle_bindings()
    )
    statement_graph = build_statement_graph(view, graph)
    guarded = next(bundle for bundle in statement_graph.bundles if bundle.guarded)
    drilldown = drilldown_bundle(
        view,
        graph,
        consumer_id=guarded.consumer_id,
        producer_id=guarded.producer_id,
        access=guarded.access,
        distance=guarded.distance,
        coeff=guarded.coeff,
        offset=guarded.offset,
        guarded=True,
    )
    assert drilldown.edges
    assert all(edge.guarded for edge in drilldown.edges)
