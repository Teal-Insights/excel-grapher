"""Statement-graph catalog facade, quotient, payload, and HTML."""

from __future__ import annotations

import inspect
import json
from pathlib import Path

import pytest

from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.exporter.semantic_catalog import load_semantic_catalog
from excel_grapher.exporter.semantic_graph import (
    REMAINDER_STATEMENT_ID,
    build_statement_graph,
)
from excel_grapher.exporter.semantic_viz import (
    SEMANTIC_VIZ_CELL_SAMPLE,
    SEMANTIC_VIZ_PAYLOAD_VERSION,
    SEMANTIC_VIZ_SVG_MAX_PRIMITIVES,
    semantic_viz_renderer,
    semantic_viz_svg_primitive_count,
    serialize_semantic_viz_json,
    spread_rank_centers,
    to_semantic_viz_payload,
    write_semantic_viz_html,
)
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
from tests.unit.exporter.inverted_tree.test_shape_a36_vintage_residual import (
    vintage_residual_bindings,
    vintage_residual_workbook,
)


def _view(workbook: Path, document: dict):
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    bindings = validate_bindings_document(document)
    view = load_semantic_catalog(graph, bindings, workbook=workbook)
    return view, graph, catalog


def test_facade_does_not_import_emit() -> None:
    import excel_grapher.exporter.semantic_catalog as facade

    source = inspect.getsource(facade)
    assert "from excel_grapher.exporter.inverted_tree.emit" not in source
    assert "from excel_grapher.exporter.inverted_tree.ast_emit" not in source
    assert "generate_inverted_tree_modules" not in source


def test_cell_web_viz_payload_does_not_require_bindings() -> None:
    params = inspect.signature(to_web_viz_payload).parameters
    assert "bindings" not in params
    assert "series_bindings" not in params


def test_facade_matches_catalog_statements(tmp_path: Path) -> None:
    view, graph, catalog = _view(_a12_workbook(tmp_path), _a12_bindings())
    path = view.catalog.get("path")
    assert len(path.statements) == 3
    assert all(stmt.cells for stmt in path.statements)
    assert view.edges.edges


def test_a12_splits_into_one_node_per_shape_run(tmp_path: Path) -> None:
    view, graph, _catalog = _view(_a12_workbook(tmp_path), _a12_bindings())
    statement_graph = build_statement_graph(view, graph)
    path_nodes = [n for n in statement_graph.nodes if n.series_id == "path" and not n.is_remainder]
    assert len(path_nodes) == 3
    assert statement_graph.stats.bundle_count < statement_graph.stats.instance_edge_count or (
        statement_graph.stats.instance_edge_count <= 2
    )
    path_ids = {n.statement_id for n in path_nodes}
    between = [
        b
        for b in statement_graph.bundles
        if b.consumer_id in path_ids and b.producer_id in path_ids
    ]
    assert between


def test_zipper_keeps_series_nodes_and_types_lag_vs_identity(tmp_path: Path) -> None:
    view, graph, _catalog = _view(_zipper_workbook(tmp_path), _zipper_bindings())
    statement_graph = build_statement_graph(view, graph)
    series_nodes = {n.series_id for n in statement_graph.nodes if not n.is_remainder}
    assert series_nodes >= {"debt", "adjustment"}
    by_series = {n.statement_id: n.series_id for n in statement_graph.nodes}
    keys = {
        (by_series[b.consumer_id], by_series[b.producer_id], b.access, b.distance)
        for b in statement_graph.bundles
        if b.consumer_id in by_series and b.producer_id in by_series
    }
    assert ("debt", "adjustment", "identity", 0) in keys
    assert any(k[0] == "adjustment" and k[1] == "debt" and k[2] == "shift" for k in keys)
    assert statement_graph.stats.bundle_count < statement_graph.stats.instance_edge_count
    identity = next(
        b
        for b in statement_graph.bundles
        if b.access == "identity" and by_series[b.consumer_id] == "debt"
    )
    by_id = {n.statement_id: i for i, n in enumerate(statement_graph.nodes)}
    assert (
        statement_graph.ranks[by_id[identity.consumer_id]]
        < statement_graph.ranks[by_id[identity.producer_id]]
    )


def test_matrix_zipper_does_not_explode_partitions(tmp_path: Path) -> None:
    view, graph, _catalog = _view(_a20_zipper_workbook(tmp_path), _a20_zipper_bindings())
    statement_graph = build_statement_graph(view, graph)
    series_nodes = [n for n in statement_graph.nodes if n.series_id in {"debt", "adjustment"}]
    assert {n.series_id for n in series_nodes} == {"debt", "adjustment"}
    assert statement_graph.stats.statement_count < statement_graph.stats.cell_count
    assert not any("France" in n.statement_id or "Kenya" in n.statement_id for n in series_nodes)
    parts = {part[0] for b in statement_graph.bundles for part in b.partitions if part}
    assert parts == {"France", "Kenya"}


def test_vintage_flags_heterogeneous_partitions(tmp_path: Path) -> None:
    view, graph, _catalog = _view(vintage_residual_workbook(tmp_path), vintage_residual_bindings())
    statement_graph = build_statement_graph(view, graph)
    series_ids = {n.series_id for n in statement_graph.nodes if not n.is_remainder}
    assert "stock" in series_ids and "principal" in series_ids
    assert not any("2024" in n.statement_id for n in statement_graph.nodes)
    assert statement_graph.stats.heterogeneous_partition_pair_count >= 1
    assert any({a, b} == {"stock", "principal"} for a, b in statement_graph.heterogeneous_pairs)


def test_guarded_access_is_a_separate_bundle(tmp_path: Path) -> None:
    view, graph, _catalog = _view(
        _series_may_cycle_workbook(tmp_path), _series_may_cycle_bindings()
    )
    statement_graph = build_statement_graph(view, graph)
    assert any(b.guarded for b in statement_graph.bundles)


def test_payload_traces_cells_and_writes_html(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    assert payload.version == SEMANTIC_VIZ_PAYLOAD_VERSION
    data = payload.to_dict()
    assert data["kind"] == "statement_graph"
    assert data["stats"]["bundle_count"] == payload.graph.stats.bundle_count
    for node in data["nodes"]:
        if not node["is_remainder"]:
            assert node["cells"]
    for bundle in data["bundles"]:
        assert bundle["representative_consumer_cell"]
        assert bundle["representative_producer_cell"]
    html_path = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html_path)
    html = html_path.read_text(encoding="utf-8")
    assert "debt" in html
    assert "statement_graph" in html or "adjustment" in html
    assert "marker-end" in html
    assert "orient', 'auto'" in html or 'orient", "auto"' in html
    assert "auto-start-reverse" not in html
    assert "stroke-opacity" in html
    assert "addEventListener('wheel'" in html or 'addEventListener("wheel"' in html
    assert 'data-dir="constant"' in html
    assert "layoutByRank" in html
    assert "cameraFor" in html


def test_html_payload_samples_cell_addresses(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    full = payload.to_dict()
    sampled = payload.to_dict(cell_sample=1)
    assert any(len(node["cells"]) > 1 for node in full["nodes"])
    for full_node, sampled_node in zip(full["nodes"], sampled["nodes"], strict=True):
        assert sampled_node["cell_count"] == full_node["cell_count"]
        assert len(sampled_node["cells"]) <= 1
        if full_node["cells"]:
            assert sampled_node["cells"] == full_node["cells"][:1]
    embedded = serialize_semantic_viz_json(payload)
    html_path = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html_path)
    html = html_path.read_text(encoding="utf-8")
    assert embedded in html
    for node in json.loads(embedded)["nodes"]:
        assert len(node["cells"]) <= SEMANTIC_VIZ_CELL_SAMPLE


def test_semantic_catalog_honors_blank_ranges(tmp_path: Path) -> None:
    from excel_grapher.exporter.semantic_catalog import SemanticCatalogError
    from excel_grapher.grapher import create_dependency_graph
    from tests.unit.exporter.inverted_tree.test_blank_ranges import (
        _BLANK,
        _mcve_bindings,
        _mcve_workbook,
    )

    workbook = _mcve_workbook(tmp_path)
    bindings = validate_bindings_document(_mcve_bindings())
    graph = create_dependency_graph(
        workbook,
        ["Outputs!B1"],
        load_values=True,
        blank_ranges=_BLANK,
    )
    with pytest.raises(SemanticCatalogError, match="is not a bound series"):
        load_semantic_catalog(graph, bindings, workbook=workbook)
    view = load_semantic_catalog(graph, bindings, workbook=workbook, blank_ranges=_BLANK)
    assert view.edges.edges
    payload = to_semantic_viz_payload(graph, bindings, workbook=workbook, blank_ranges=_BLANK)
    assert payload.graph.stats.statement_count >= 1


def test_remainder_node_when_graph_has_unbound_cells(tmp_path: Path) -> None:
    view, graph, catalog = _view(_zipper_workbook(tmp_path), _zipper_bindings())
    from excel_grapher.grapher.node import make_cell_node

    graph.add_node(make_cell_node("Engine", "Z", 99, value=1, is_leaf=True))
    statement_graph = build_statement_graph(view, graph)
    remainder = [n for n in statement_graph.nodes if n.is_remainder]
    assert remainder
    assert remainder[0].statement_id == REMAINDER_STATEMENT_ID
    assert statement_graph.stats.unbound_cell_count >= 1
    assert "Engine!Z99" in catalog.address_to_id or "Engine!Z99" in statement_graph.remainder_sample


def test_qcraft_scale_stays_on_svg() -> None:
    # Q-CRAFT: 715 statements / 5,725 bundles. SVG first paint was usable.
    count = semantic_viz_svg_primitive_count(statement_count=715, bundle_count=5725)
    assert count < SEMANTIC_VIZ_SVG_MAX_PRIMITIVES
    assert semantic_viz_renderer(statement_count=715, bundle_count=5725) == "svg"


def test_lic_dsf_scale_uses_canvas() -> None:
    # LIC-DSF statement grain: ~20k statements / ~120k bundles. SVG paint hung.
    assert semantic_viz_svg_primitive_count(statement_count=20020, bundle_count=119808) > (
        SEMANTIC_VIZ_SVG_MAX_PRIMITIVES
    )
    assert semantic_viz_renderer(statement_count=20020, bundle_count=119808) == "canvas"


def test_spread_rank_centers_fills_camera_not_left_edge() -> None:
    widths = (80.0, 80.0, 80.0)
    centers = spread_rank_centers(
        widths, pad=48, min_gap=16, min_row_width=2800, graph_width=5_000_000
    )
    assert centers[0] >= 48
    assert centers[-1] <= 2800 - 48
    assert centers[-1] - centers[0] > 2000


def test_spread_rank_centers_keeps_dense_packing() -> None:
    widths = tuple(72.0 for _ in range(40))
    packed_span = 40 * 72 + 39 * 16
    centers = spread_rank_centers(
        widths, pad=48, min_gap=16, min_row_width=2800, graph_width=packed_span + 96
    )
    gaps = [centers[i + 1] - centers[i] - 72 for i in range(len(centers) - 1)]
    assert all(abs(g - 16) < 1e-6 for g in gaps)


def test_html_ships_canvas_painter_and_rank_spread(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    html_path = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html_path)
    html = html_path.read_text(encoding="utf-8")
    assert "getContext('2d')" in html
    assert "spread_rank_centers" in html or "spreadRank" in html
    assert str(SEMANTIC_VIZ_SVG_MAX_PRIMITIVES) in html
    assert "roundRect" in html or "quadraticCurveTo" in html
    assert payload.graph.stats.statement_count + 2 * payload.graph.stats.bundle_count < (
        SEMANTIC_VIZ_SVG_MAX_PRIMITIVES
    )
