from __future__ import annotations

import math

import pytest

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import Literal
from excel_grapher.grapher.lightweight_viz import (
    MODULE_INFERENCE_OVERLAY_ID,
    VIZ_PAYLOAD_VERSION,
    VizLimits,
    assemble_lightweight_viz_payload,
    build_lightweight_viz_core,
    lightweight_viz_flat,
)
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.overlay_registry import build_overlays


def _tiny_graph() -> DependencyGraph:
    g = DependencyGraph()
    n = Node(
        sheet="S",
        column="A",
        row=1,
        formula=None,
        normalized_formula=None,
        value=1,
        is_leaf=True,
    )
    g.add_node(n)
    return g


def _tiny_chain_graph() -> DependencyGraph:
    g = DependencyGraph()
    n1 = Node(
        sheet="S",
        column="A",
        row=1,
        formula=None,
        normalized_formula=None,
        value=1,
        is_leaf=True,
    )
    n2 = Node(
        sheet="S",
        column="A",
        row=2,
        formula="=A1",
        normalized_formula="=A1",
        value=None,
        is_leaf=False,
    )
    n3 = Node(
        sheet="S",
        column="A",
        row=3,
        formula="=A2",
        normalized_formula="=A2",
        value=None,
        is_leaf=False,
    )
    for n in (n1, n2, n3):
        g.add_node(n)
    g.add_edge(n3.key, n2.key)
    g.add_edge(n2.key, n1.key)
    return g


def _two_component_graph() -> DependencyGraph:
    g = DependencyGraph()
    a1 = Node(
        sheet="S",
        column="A",
        row=1,
        formula=None,
        normalized_formula=None,
        value=1,
        is_leaf=True,
    )
    a2 = Node(
        sheet="S",
        column="A",
        row=2,
        formula="=A1",
        normalized_formula="=A1",
        value=None,
        is_leaf=False,
    )
    b1 = Node(
        sheet="S",
        column="B",
        row=1,
        formula=None,
        normalized_formula=None,
        value=1,
        is_leaf=True,
    )
    b2 = Node(
        sheet="S",
        column="B",
        row=2,
        formula="=B1",
        normalized_formula="=B1",
        value=None,
        is_leaf=False,
    )
    for n in (a1, a2, b1, b2):
        g.add_node(n)
    g.add_edge(a2.key, a1.key)
    g.add_edge(b2.key, b1.key)
    return g


def test_register_overlay_duplicate_rejected() -> None:
    from excel_grapher.exporter.lightweight_viz import ensure_default_overlay_builders
    from excel_grapher.grapher.lightweight_viz import LightweightVizOverlay
    from excel_grapher.grapher.overlay_registry import (
        clear_overlay_builders_for_tests,
        register_overlay_builder,
    )

    clear_overlay_builders_for_tests()

    def _dummy(_g, _c, *, context) -> LightweightVizOverlay:
        return LightweightVizOverlay(
            overlay_id="tests.dummy",
            schema_version=1,
            kind="test",
            data={"x": 1},
        )

    oid = "tests.overlay_registry_duplicate"
    register_overlay_builder(oid, _dummy)
    with pytest.raises(ValueError, match="duplicate overlay_id"):
        register_overlay_builder(oid, _dummy)

    ensure_default_overlay_builders()


def test_build_overlays_builder_failure_includes_overlay_id() -> None:
    from excel_grapher.grapher.lightweight_viz import LightweightVizOverlay
    from excel_grapher.grapher.overlay_registry import (
        build_overlays,
        clear_overlay_builders_for_tests,
        register_overlay_builder,
    )

    clear_overlay_builders_for_tests()

    oid = "tests.overlay_raises"

    def _raises(_graph, _core, *, context) -> LightweightVizOverlay:
        raise ValueError("simulated builder failure")

    register_overlay_builder(oid, _raises)
    core = build_lightweight_viz_core(_tiny_graph(), limits=VizLimits())
    with pytest.raises(RuntimeError, match=r"overlay builder failed for overlay_id='tests\.overlay_raises'"):
        build_overlays(_tiny_graph(), core, [oid])

    from excel_grapher.exporter.lightweight_viz import ensure_default_overlay_builders

    ensure_default_overlay_builders()


def test_build_overlays_unknown_id_errors() -> None:
    from excel_grapher.exporter.lightweight_viz import ensure_default_overlay_builders

    ensure_default_overlay_builders()
    core = build_lightweight_viz_core(_tiny_graph(), limits=VizLimits())
    with pytest.raises(ValueError, match="unknown overlay_id"):
        build_overlays(_tiny_graph(), core, ["no.such.overlay"])


def test_core_only_flat_shape() -> None:
    core = build_lightweight_viz_core(_tiny_graph(), limits=VizLimits(), layout_input=None)
    wire_payload = assemble_lightweight_viz_payload(core, [])
    flat = lightweight_viz_flat(wire_payload)
    assert wire_payload.version == VIZ_PAYLOAD_VERSION
    assert flat.stats.node_count == 1
    assert flat.stats.module_count == 1
    assert flat.stats.scc_count == 0


def test_module_overlay_id_constant_matches_exporter() -> None:
    assert MODULE_INFERENCE_OVERLAY_ID == "exporter.module_inference"


def test_core_default_layout_is_bfs_from_targets() -> None:
    g = _tiny_chain_graph()
    core = build_lightweight_viz_core(g, limits=VizLimits())
    keys = sorted(g)
    idx = {k: i for i, k in enumerate(keys)}
    ranks = core.nodes.rank
    assert ranks[idx["S!A3"]] <= ranks[idx["S!A2"]] <= ranks[idx["S!A1"]]


def test_core_layout_modes_available() -> None:
    g = _tiny_chain_graph()
    bfs = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="bfs")
    layered = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="layered")
    grid = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="grid")
    assert bfs.stats.node_count == layered.stats.node_count == grid.stats.node_count == 3


def test_core_force_layout_accepted_and_positions_valid() -> None:
    g = _tiny_chain_graph()
    core = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="force")
    assert core.stats.node_count == 3
    assert core.stats.local_edge_count == 2
    xs = list(core.nodes.x)
    ys = list(core.nodes.y)
    assert all(math.isfinite(x) and math.isfinite(y) for x, y in zip(xs, ys, strict=True))
    assert not all(x == xs[0] for x in xs) or not all(y == ys[0] for y in ys)


def test_core_force_layout_deterministic() -> None:
    g = _tiny_chain_graph()
    a = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="force")
    b = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="force")
    assert list(a.nodes.x) == list(b.nodes.x)
    assert list(a.nodes.y) == list(b.nodes.y)


def test_core_force_layout_differs_from_rank_band_on_chain() -> None:
    """Force-mode coordinates differ from rank-band layouts on the tiny chain fixture."""
    g = _tiny_chain_graph()
    bfs = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="bfs")
    layered = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="layered")
    grid = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="grid")
    force = build_lightweight_viz_core(g, limits=VizLimits(), layout_mode="force")
    assert list(bfs.nodes.rank) == list(layered.nodes.rank)
    assert list(bfs.nodes.x) != list(force.nodes.x) or list(bfs.nodes.y) != list(force.nodes.y)
    assert all(math.isfinite(x) for x in bfs.nodes.x)
    assert all(math.isfinite(x) for x in layered.nodes.x)
    assert all(math.isfinite(x) for x in grid.nodes.x)


def test_core_bfs_can_exclude_unreachable_from_explicit_seed_set() -> None:
    g = _two_component_graph()
    core = build_lightweight_viz_core(
        g,
        limits=VizLimits(),
        layout_mode="bfs",
        bfs_seed_keys=("S!A2",),
        exclude_unreachable_from_bfs=True,
    )
    assert core.stats.node_count == 2


def test_core_excluding_guarded_edges_prunes_unreachable_nodes() -> None:
    g = DependencyGraph()
    a = Node(
        sheet="S",
        column="A",
        row=1,
        formula="=B1",
        normalized_formula="=B1",
        value=None,
        is_leaf=False,
    )
    b = Node(
        sheet="S",
        column="B",
        row=1,
        formula=None,
        normalized_formula=None,
        value=1,
        is_leaf=True,
    )
    c = Node(
        sheet="S",
        column="C",
        row=1,
        formula=None,
        normalized_formula=None,
        value=1,
        is_leaf=True,
    )
    for n in (a, b, c):
        g.add_node(n)
    g.add_edge(a.key, b.key)
    g.add_edge(a.key, c.key, guard=Literal(True))

    all_edges_core = build_lightweight_viz_core(
        g, limits=VizLimits(), layout_mode="bfs", bfs_seed_keys=(a.key,)
    )
    unguarded_core = build_lightweight_viz_core(
        g,
        limits=VizLimits(),
        layout_mode="bfs",
        include_guarded_edges=False,
        bfs_seed_keys=(a.key,),
    )
    assert all_edges_core.stats.node_count == 3
    assert unguarded_core.stats.node_count == 2
