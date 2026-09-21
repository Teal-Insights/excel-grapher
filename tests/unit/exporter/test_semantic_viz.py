"""Statement-graph catalog facade, quotient, payload, and HTML."""

from __future__ import annotations

import inspect
import json
import math
import shutil
import subprocess
from pathlib import Path

import pytest

from excel_grapher.core.address_keys import canonical_address
from excel_grapher.exporter import semantic_viz as semantic_viz_mod
from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.exporter.semantic_catalog import load_semantic_catalog
from excel_grapher.exporter.semantic_graph import (
    MIXED_SHEET,
    REMAINDER_STATEMENT_ID,
    build_statement_graph,
    statement_sheet,
)
from excel_grapher.exporter.semantic_viz import (
    SEMANTIC_VIZ_BOX_MAX_PRIMITIVES,
    SEMANTIC_VIZ_CELL_SAMPLE,
    SEMANTIC_VIZ_PAYLOAD_VERSION,
    semantic_viz_clustered_layout_allowed,
    semantic_viz_primitive_count,
    serialize_semantic_viz_json,
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


def test_statement_sheet_is_mixed_when_cells_span_sheets() -> None:
    same = (canonical_address("Engine!A2"), canonical_address("Engine!B2"))
    assert statement_sheet(same) == "Engine"
    quoted = (canonical_address("'My Sheet'!A1"),)
    assert statement_sheet(quoted) == "My Sheet"
    mixed = (canonical_address("Engine!A2"), canonical_address("Other!A1"))
    assert statement_sheet(mixed) == MIXED_SHEET
    assert statement_sheet(()) == MIXED_SHEET


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
            assert node["sheet"] == "Engine"
        else:
            assert node["sheet"]
    for bundle in data["bundles"]:
        assert bundle["representative_consumer_cell"]
        assert bundle["representative_producer_cell"]
    html_path = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html_path)
    html = html_path.read_text(encoding="utf-8")
    assert "debt" in html
    assert "statement_graph" in html or "adjustment" in html
    assert "canvas" in html.lower()


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


def test_payload_traces_nodes_and_bundles(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    full = payload.to_dict()
    sampled = payload.to_dict(cell_sample=1)
    bound_nodes = [node for node in full["nodes"] if not node["is_remainder"]]
    assert bound_nodes
    for node in bound_nodes:
        assert node["cells"]
        assert node["cell_count"] == len(node["cells"])
        assert "partitions" in node
        assert node["cell_partitions"] == [[] for _ in node["cells"]]
    for bundle in full["bundles"]:
        assert bundle["instance_edges"]
        assert len(bundle["instance_edges"]) == bundle["instance_edge_count"]
        first = bundle["instance_edges"][0]
        assert first["consumer_cell"] == bundle["representative_consumer_cell"]
        assert first["producer_cell"] == bundle["representative_producer_cell"]
    for full_bundle, sampled_bundle in zip(full["bundles"], sampled["bundles"], strict=True):
        assert sampled_bundle["instance_edge_count"] == full_bundle["instance_edge_count"]
        assert len(sampled_bundle["instance_edges"]) <= 1
        if full_bundle["instance_edges"]:
            assert sampled_bundle["instance_edges"] == full_bundle["instance_edges"][:1]
    embedded = json.loads(serialize_semantic_viz_json(payload))
    for bundle in embedded["bundles"]:
        assert len(bundle["instance_edges"]) <= SEMANTIC_VIZ_CELL_SAMPLE


def test_html_ships_inspect_panel(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    html = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html)
    text = html.read_text(encoding="utf-8")
    assert 'id="inspect"' in text
    assert 'id="inspectTitle"' in text
    assert 'id="inspectBody"' in text
    assert "function showInspect" in text
    assert "function hitNode" in text
    assert "function hitBundle" in text
    assert "instance_edges" in text
    assert "cell_partitions" in text
    assert "drilldown" in text.lower() or "Inspect" in text


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
    graph.add_node(make_cell_node("Other", "A", 1, value=1, is_leaf=True))
    statement_graph = build_statement_graph(view, graph)
    remainder = [n for n in statement_graph.nodes if n.is_remainder]
    assert remainder
    assert remainder[0].statement_id == REMAINDER_STATEMENT_ID
    assert remainder[0].sheet == MIXED_SHEET
    assert statement_graph.stats.unbound_cell_count >= 2
    assert "Engine!Z99" in catalog.address_to_id or "Engine!Z99" in statement_graph.remainder_sample
    payload = to_semantic_viz_payload(
        graph,
        validate_bindings_document(_zipper_bindings()),
        workbook=_zipper_workbook(tmp_path),
        view=view,
    )
    remainder_payload = next(node for node in payload.to_dict()["nodes"] if node["is_remainder"])
    assert remainder_payload["sheet"] == MIXED_SHEET


def test_qcraft_scale_uses_boxes() -> None:
    # Q-CRAFT: 715 statements / 5,725 bundles. Labeled boxes stay readable.
    count = semantic_viz_primitive_count(statement_count=715, bundle_count=5725)
    assert count < SEMANTIC_VIZ_BOX_MAX_PRIMITIVES


def test_lic_dsf_scale_uses_dots() -> None:
    # LIC-DSF statement grain: ~20k statements / ~120k bundles.
    assert semantic_viz_primitive_count(statement_count=20020, bundle_count=119808) > (
        SEMANTIC_VIZ_BOX_MAX_PRIMITIVES
    )


def test_html_ships_canvas_painter_and_rank_spread(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    html_path = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html_path)
    html = html_path.read_text(encoding="utf-8")
    assert "canvas" in html.lower()
    assert str(SEMANTIC_VIZ_BOX_MAX_PRIMITIVES) in html
    assert payload.graph.stats.statement_count + 2 * payload.graph.stats.bundle_count < (
        SEMANTIC_VIZ_BOX_MAX_PRIMITIVES
    )


def test_html_ships_legend_and_reset_control(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    html = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html)
    text = html.read_text(encoding="utf-8")

    assert 'id="legend"' in text
    assert 'style="color:#0969da"' in text
    assert 'style="color:#8250df"' in text
    assert 'style="color:#1a7f37"' in text
    assert 'style="color:#9a6700"' in text
    assert "identity: '#0969da'" in text
    assert "shift: '#8250df'" in text
    assert "affine: '#1a7f37'" in text
    assert "gather: '#9a6700'" in text
    assert "whole: '#9a6700'" in text
    assert "dynamic: '#9a6700'" in text
    assert "cross_partition: '#9a6700'" in text
    assert "lag" in text.lower()
    assert "dashed" in text.lower()
    assert "contemporaneous" in text.lower()
    assert "solid" in text.lower()
    assert "guarded" in text.lower()
    assert "per-series hue" in text.lower()
    assert "unbound remainder" in text.lower()
    assert "#d0d7de" in text

    assert 'id="reset"' in text
    assert ">Reset</button>" in text
    assert 'getElementById("reset")' in text or "getElementById('reset')" in text
    assert "draw(true)" in text
    assert ">Fit</button>" not in text

    assert 'id="search"' in text
    for direction in ("constant", "input", "internal", "output"):
        assert f'data-dir="{direction}"' in text
    assert "pointermove" in text
    assert "wheel" in text


def test_html_ships_color_and_cluster_controls(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    html = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html)
    text = html.read_text(encoding="utf-8")

    assert 'id="colorBy"' in text
    assert 'id="clusterBy"' in text
    assert 'value="series"' in text
    assert 'value="role"' in text
    assert 'value="sheet"' in text
    assert 'value="none"' in text
    assert 'id="colorBy"' in text and "selected" in text
    assert 'option value="series" selected' in text or 'value="series" selected' in text
    assert 'value="none"' in text
    assert 'id="clusterBy"' in text
    cluster_block = text[text.index('id="clusterBy"') : text.index('id="clusterBy"') + 400]
    assert 'value="series" selected' in cluster_block
    assert 'value="none" selected' not in cluster_block

    assert "ROLE_COLORS" in text
    assert "constant:" in text
    assert "input:" in text
    assert "internal:" in text
    assert "output:" in text
    assert "remainder:" in text
    assert "function nodeFill" in text
    assert "function clusterKey" in text
    assert "function paintClusters" in text
    assert "function updateLegend" in text
    assert "MIXED_SHEET" in text or "'mixed'" in text or '"mixed"' in text
    assert "colorMode" in text
    assert "clusterMode" in text
    assert "getElementById('colorBy').addEventListener" in text or (
        'getElementById("colorBy").addEventListener' in text
    )
    assert "updateLegend();" in text
    assert "paintCanvas();" in text
    assert "getElementById('clusterBy').addEventListener" in text or (
        'getElementById("clusterBy").addEventListener' in text
    )
    assert "draw(false)" in text
    assert "clusterMode() === 'none'" in text or 'clusterMode() === "none"' in text
    assert "n.sheet" in text
    assert "n.direction" in text
    assert "identity: '#0969da'" in text
    assert 'id="legendFill"' in text
    assert 'id="search"' in text
    for direction in ("constant", "input", "internal", "output"):
        assert f'data-dir="{direction}"' in text
    assert 'id="reset"' in text
    assert "pointermove" in text
    assert "wheel" in text


def _layout_js_path() -> Path:
    return Path(semantic_viz_mod.__file__).with_name("semantic_viz_layout.js")


def _node_bin() -> str:
    exe = shutil.which("node")
    if exe is None:
        pytest.skip("node is required to execute statement-graph layout JS")
    return exe


def _run_layout_js(payload: dict) -> dict:
    """Execute `semantic_viz_layout.js` against a JSON payload."""
    runner = """
    const fs = require("fs");
    const api = require(process.argv[process.argv.length - 1]);
    const input = JSON.parse(fs.readFileSync(0, "utf8"));
    const nodes = (input.nodes || []).map((n) => Object.assign({}, n));
    nodes.forEach((n) => {
      n._hidden = !!n._hidden;
      n._label = n._label || n.id;
      n._cluster = api.clusterKey(n, input.clusterBy || "none");
    });
    const size = api.layoutClustered(nodes, input.bundles || [], !!input.compact);
    process.stdout.write(JSON.stringify({
      size: size,
      allowed: api.clusteredLayoutAllowed(input.primitiveCount, input.boxMaxPrimitives),
      positions: Object.fromEntries(nodes.map((n) => [n.id, {x: n._x, y: n._y}])),
      hulls: api.clusterHulls(nodes, {splitByRank: false}).map((h) => ({
        key: h.key, count: h.members.length
      })),
      rankHulls: api.clusterHulls(nodes, {splitByRank: true}).map((h) => ({
        key: h.key, rank: h.rank, count: h.members.length
      }))
    }));
    """
    proc = subprocess.run(
        [_node_bin(), "-e", runner, str(_layout_js_path())],
        input=json.dumps(payload),
        capture_output=True,
        text=True,
        check=False,
    )
    if proc.returncode != 0:
        raise AssertionError(proc.stderr or proc.stdout)
    return json.loads(proc.stdout)


def test_clustered_layout_allowed_below_box_cap_only() -> None:
    assert semantic_viz_clustered_layout_allowed(statement_count=715, bundle_count=5725)
    assert semantic_viz_clustered_layout_allowed(statement_count=25, bundle_count=40)
    assert not semantic_viz_clustered_layout_allowed(statement_count=20020, bundle_count=119808)


def test_clustered_layout_moves_nodes_when_cluster_by_changes(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    data = payload.to_dict()
    ids = [node["id"] for node in data["nodes"]]
    assert ids
    base = {"nodes": data["nodes"], "bundles": data["bundles"]}
    none = _run_layout_js({**base, "clusterBy": "none"})
    series = _run_layout_js({**base, "clusterBy": "series"})
    role = _run_layout_js({**base, "clusterBy": "role"})

    def moved(left: dict, right: dict) -> list[str]:
        return [
            node_id
            for node_id in ids
            if math.hypot(
                left["positions"][node_id]["x"] - right["positions"][node_id]["x"],
                left["positions"][node_id]["y"] - right["positions"][node_id]["y"],
            )
            > 1.0
        ]

    assert moved(none, series)
    assert moved(series, role)
    assert none["hulls"] == []
    series_keys = {hull["key"] for hull in series["hulls"]}
    assert series_keys >= {"debt", "adjustment"}
    assert len(series["hulls"]) == len(series_keys)
    assert len(series["rankHulls"]) >= len(series["hulls"])


def test_clustered_sheet_layout_keeps_mixed_as_own_group() -> None:
    nodes = [
        {
            "id": "eng",
            "series_id": "alpha",
            "direction": "internal",
            "sheet": "Engine",
            "start": 0,
            "is_remainder": False,
            "rank": 0,
        },
        {
            "id": "mix",
            "series_id": "beta",
            "direction": "internal",
            "sheet": MIXED_SHEET,
            "start": 0,
            "is_remainder": False,
            "rank": 1,
        },
        {
            "id": "oth",
            "series_id": "gamma",
            "direction": "output",
            "sheet": "Other",
            "start": 0,
            "is_remainder": False,
            "rank": 0,
        },
        {
            "id": "eng2",
            "series_id": "alpha",
            "direction": "internal",
            "sheet": "Engine",
            "start": 1,
            "is_remainder": False,
            "rank": 2,
        },
    ]
    out = _run_layout_js({"nodes": nodes, "bundles": [], "clusterBy": "sheet"})
    keys = {hull["key"] for hull in out["hulls"]}
    assert keys == {"Engine", MIXED_SHEET, "Other"}
    mix = out["positions"]["mix"]
    engine = out["positions"]["eng"]
    assert math.hypot(mix["x"] - engine["x"], mix["y"] - engine["y"]) > 20
    rank_keys = {(hull["key"], hull["rank"]) for hull in out["rankHulls"]}
    assert ("Engine", 0) in rank_keys
    assert ("Engine", 2) in rank_keys
    assert len(out["hulls"]) == 3
    assert len(out["rankHulls"]) > len(out["hulls"])


def test_layout_js_refuses_force_above_box_cap() -> None:
    out = _run_layout_js(
        {
            "nodes": [],
            "bundles": [],
            "clusterBy": "none",
            "primitiveCount": semantic_viz_primitive_count(
                statement_count=20020, bundle_count=119808
            ),
            "boxMaxPrimitives": SEMANTIC_VIZ_BOX_MAX_PRIMITIVES,
        }
    )
    assert out["allowed"] is False
    small = _run_layout_js(
        {
            "nodes": [],
            "bundles": [],
            "clusterBy": "none",
            "primitiveCount": semantic_viz_primitive_count(statement_count=715, bundle_count=5725),
            "boxMaxPrimitives": SEMANTIC_VIZ_BOX_MAX_PRIMITIVES,
        }
    )
    assert small["allowed"] is True


def test_html_ships_clustered_layout_mode(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    view, graph, _catalog = _view(workbook, _zipper_bindings())
    payload = to_semantic_viz_payload(
        graph, validate_bindings_document(_zipper_bindings()), workbook=workbook, view=view
    )
    html = tmp_path / "zipper.html"
    write_semantic_viz_html(payload, html)
    text = html.read_text(encoding="utf-8")

    assert 'id="layoutMode"' in text
    assert 'value="rank" selected' in text or 'option value="rank" selected' in text
    assert 'value="clustered"' in text
    assert "layoutClustered" in text
    assert "clusteredLayoutAllowed" in text
    assert "clusterHulls" in text
    assert "splitByRank" in text
    assert "SemanticVizLayout" in text
    assert "markStyle === 'dots'" in text or 'markStyle === "dots"' in text
    assert "clusteredOpt.disabled" in text or "clustered.disabled" in text
    assert "layoutMode() === 'clustered'" in text or 'layoutMode() === "clustered"' in text
    assert "getElementById('layoutMode')" in text or 'getElementById("layoutMode")' in text
    assert "canvas-side" in text.lower() or "clustered force" in text.lower()
    assert 'id="colorBy"' in text
    assert 'id="clusterBy"' in text
    assert 'id="legendFill"' in text
    assert "identity: '#0969da'" in text
    assert 'id="reset"' in text
    assert 'id="search"' in text
    for direction in ("constant", "input", "internal", "output"):
        assert f'data-dir="{direction}"' in text
    assert "pointermove" in text
    assert "wheel" in text
    assert payload.graph.stats.statement_count + 2 * payload.graph.stats.bundle_count < (
        SEMANTIC_VIZ_BOX_MAX_PRIMITIVES
    )
