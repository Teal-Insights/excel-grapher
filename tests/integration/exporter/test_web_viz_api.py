"""Graph-first web viz API coverage (integration)."""

from __future__ import annotations

import importlib.resources
import inspect
import json
import re
from pathlib import Path

import pytest

from excel_grapher import write_web_viz_html
from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.exporter.web_viz_layout import (
    LAYOUT_CLUSTERED_FORCE,
    WebVizLayoutContext,
    WebVizLayoutResult,
    list_web_viz_layouts,
    register_web_viz_layout,
    unregister_web_viz_layout,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from tests.viz_marks import requires_numpy


def _build_two_component_digraph():
    import networkx as nx

    g = nx.DiGraph()
    g.add_node("S!A1", formula=None, value=1, is_leaf=True, sheet="S", column="A", row=1)
    g.add_node("S!A2", formula="=A1", value=None, is_leaf=False, sheet="S", column="A", row=2)
    g.add_node("S!B1", formula=None, value=1, is_leaf=True, sheet="S", column="B", row=1)
    g.add_node("S!B2", formula="=B1", value=None, is_leaf=False, sheet="S", column="B", row=2)
    g.add_edge("S!A2", "S!A1")
    g.add_edge("S!B2", "S!B1")
    return g


def _build_chain_digraph(n: int):
    import networkx as nx

    g = nx.DiGraph()
    for r in range(1, n + 1):
        g.add_node(
            f"S!A{r}",
            formula=None if r == 1 else f"=A{r - 1}",
            is_leaf=(r == 1),
            value=1 if r == 1 else None,
        )
    for r in range(2, n + 1):
        g.add_edge(f"S!A{r}", f"S!A{r - 1}")
    return g


def _build_two_component_graph() -> DependencyGraph:
    graph = DependencyGraph()
    for col, rows in (("A", (1, 2)), ("B", (1, 2))):
        graph.add_node(
            Node(
                sheet="S",
                column=col,
                row=rows[0],
                formula=None,
                value=1,
                is_leaf=True,
            )
        )
        graph.add_node(
            Node(
                sheet="S",
                column=col,
                row=rows[1],
                formula=f"={col}1",
                value=None,
                is_leaf=False,
            )
        )
        graph.add_edge(f"S!{col}2", f"S!{col}1")
    return graph


def _zero_layout(ctx: WebVizLayoutContext, layout_config: dict[str, object]) -> WebVizLayoutResult:
    del layout_config
    return WebVizLayoutResult(
        positions={key: (0.0, 0.0) for key in ctx.keys},
        module_analysis=None,
        annotations={"custom_layout": "direct"},
        viewer_hints={},
    )


def test_to_web_viz_payload_default_layout_is_clustered_force() -> None:
    sig = inspect.signature(to_web_viz_payload)
    assert sig.parameters["layout"].default == "clustered_force"


def test_to_web_viz_layout_registry_includes_builtins() -> None:
    ids = set(list_web_viz_layouts())
    assert "clustered_force" in ids
    assert "spring" in ids
    assert "stratified_multipartite" not in ids
    assert "multipartite" not in ids


def test_register_web_viz_layout_allows_replace() -> None:
    layout_id = "_test_replace_layout"

    def first(ctx: WebVizLayoutContext, layout_config: dict[str, object]) -> WebVizLayoutResult:
        del layout_config
        return WebVizLayoutResult(
            positions={key: (0.0, 0.0) for key in ctx.keys},
            module_analysis=None,
            annotations={"custom_layout": "first"},
            viewer_hints={},
        )

    def second(ctx: WebVizLayoutContext, layout_config: dict[str, object]) -> WebVizLayoutResult:
        del layout_config
        return WebVizLayoutResult(
            positions={key: (1.0, 1.0) for key in ctx.keys},
            module_analysis=None,
            annotations={"custom_layout": "second"},
            viewer_hints={},
        )

    try:
        register_web_viz_layout(layout_id, first)
        with pytest.raises(ValueError, match="duplicate"):
            register_web_viz_layout(layout_id, first)
        register_web_viz_layout(layout_id, second, replace=True)

        payload = to_web_viz_payload(_build_two_component_digraph(), layout=layout_id)
        assert payload.annotations is not None
        assert payload.annotations["custom_layout"] == "second"
    finally:
        unregister_web_viz_layout(layout_id)


def test_to_web_viz_payload_accepts_direct_layout_plugin() -> None:
    def custom_layout(
        ctx: WebVizLayoutContext, layout_config: dict[str, object]
    ) -> WebVizLayoutResult:
        del layout_config
        return WebVizLayoutResult(
            positions={key: (0.5, 0.5) for key in ctx.keys},
            module_analysis=None,
            annotations={"custom_layout": "direct"},
            viewer_hints={},
        )

    payload = to_web_viz_payload(_build_two_component_digraph(), layout=custom_layout)
    assert payload.annotations is not None
    assert payload.annotations["custom_layout"] == "direct"


def test_to_web_viz_payload_unknown_layout_raises() -> None:
    g = _build_two_component_digraph()
    with pytest.raises(ValueError, match="Unknown web viz layout"):
        to_web_viz_payload(g, layout="not.a.registered.id", seed=0)


def test_to_web_viz_payload_rejects_unsupported_type() -> None:
    with pytest.raises(TypeError, match="DependencyGraph"):
        to_web_viz_payload(object())


def test_nx_layouts_do_not_call_to_networkx_for_graph(monkeypatch: pytest.MonkeyPatch) -> None:
    pytest.importorskip("numpy")

    def boom(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("NetworkX layouts must build a work graph from GraphReadView")

    monkeypatch.setattr("excel_grapher.grapher.export.to_networkx", boom)
    payload = to_web_viz_payload(_build_two_component_graph(), layout="spring", seed=11)
    assert payload.core.stats.node_count == 4


@requires_numpy
def test_digraph_compat_path_still_reconstructs(monkeypatch: pytest.MonkeyPatch) -> None:
    import excel_grapher.exporter.lightweight_viz as lv

    calls: list[int] = []
    real = lv._dependency_graph_from_networkx

    def counted(nx_graph: object) -> object:
        calls.append(1)
        return real(nx_graph)

    monkeypatch.setattr(lv, "_dependency_graph_from_networkx", counted)
    payload = to_web_viz_payload(_build_two_component_digraph(), seed=7)
    assert calls == [1]
    assert payload.core.stats.node_count == 4


@requires_numpy
def test_to_web_viz_payload_includes_annotations() -> None:
    g = _build_two_component_digraph()
    payload = to_web_viz_payload(g, seed=7, layout=LAYOUT_CLUSTERED_FORCE)
    assert payload.annotations is not None
    assert payload.annotations.get("layout") == "clustered_force"
    assert payload.annotations.get("rank_pull") == "between"
    assert payload.annotations.get("tier") == "small"

    from excel_grapher.grapher.lightweight_viz import serialize_lightweight_viz_json

    blob = json.loads(serialize_lightweight_viz_json(payload))
    assert blob.get("annotations", {}).get("layout") == "clustered_force"
    assert blob["version"] == 3
    assert "bucket_density" not in blob["core"]["nodes"]
    assert "dense_bucket_count" not in blob["core"]["stats"]
    assert len(blob["core"]["nodes"]["depth"]) == 4


@requires_numpy
def test_to_web_viz_payload_accepts_networkx_digraph() -> None:
    g = _build_two_component_digraph()

    payload = to_web_viz_payload(g, seed=7)
    assert payload.core.stats.node_count == 4
    assert len(payload.overlays[0].data["modules"]) == 2
    assert payload.overlays[0].display_name == "Directed Louvain modules"
    assert payload.overlays[0].overlay_id == "webviz.louvain_directed"


@requires_numpy
def test_to_web_viz_payload_accepts_dependency_graph() -> None:
    graph = _build_two_component_graph()
    payload = to_web_viz_payload(graph, seed=7)
    assert payload.core.stats.node_count == 4
    assert len(payload.overlays[0].data["modules"]) == 2
    assert payload.overlays[0].overlay_id == "webviz.louvain_directed"


@requires_numpy
def test_to_web_viz_payload_skips_nx_reconstruction_for_graph(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def boom(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("DependencyGraph must not round-trip through NetworkX")

    monkeypatch.setattr(
        "excel_grapher.exporter.lightweight_viz._dependency_graph_from_networkx",
        boom,
    )
    payload = to_web_viz_payload(_build_two_component_graph(), seed=7)
    assert payload.core.stats.node_count == 4


@requires_numpy
def test_default_layout_does_not_call_to_networkx_for_graph(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def boom(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("clustered_force must not force to_networkx()")

    monkeypatch.setattr("excel_grapher.grapher.export.to_networkx", boom)
    payload = to_web_viz_payload(_build_two_component_graph(), seed=7)
    assert payload.core.stats.node_count == 4


def test_custom_layout_does_not_materialize_networkx(monkeypatch: pytest.MonkeyPatch) -> None:
    def boom(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("custom layouts that ignore ctx.nx_graph must not convert")

    monkeypatch.setattr("excel_grapher.grapher.export.to_networkx", boom)
    payload = to_web_viz_payload(
        _build_two_component_graph(),
        layout=_zero_layout,
        include_module_overlay=False,
    )
    assert payload.annotations is not None
    assert payload.annotations["custom_layout"] == "direct"


def test_layout_plugin_can_opt_into_nx_graph(monkeypatch: pytest.MonkeyPatch) -> None:
    import excel_grapher.grapher.export as export_mod

    calls: list[int] = []
    real_to_networkx = export_mod.to_networkx

    def counted_to_networkx(graph: object, **kwargs: object) -> object:
        calls.append(1)
        return real_to_networkx(graph, **kwargs)

    monkeypatch.setattr(export_mod, "to_networkx", counted_to_networkx)

    def uses_nx(ctx: WebVizLayoutContext, layout_config: dict[str, object]) -> WebVizLayoutResult:
        del layout_config
        nx_graph = ctx.nx_graph
        assert nx_graph is not None
        return WebVizLayoutResult(
            positions={key: (0.0, 0.0) for key in ctx.keys},
            module_analysis=None,
            annotations={"nx_nodes": nx_graph.number_of_nodes()},
            viewer_hints={},
        )

    payload = to_web_viz_payload(
        _build_two_component_graph(),
        layout=uses_nx,
        include_module_overlay=False,
    )
    assert calls == [1]
    assert payload.annotations is not None
    assert payload.annotations["nx_nodes"] == 4


@requires_numpy
def test_to_web_viz_payload_is_deterministic_with_seed() -> None:
    g = _build_two_component_digraph()

    a = to_web_viz_payload(g, seed=17)
    b = to_web_viz_payload(g, seed=17)

    assert a.overlays[0].data["node_module_id"] == b.overlays[0].data["node_module_id"]
    assert a.core.nodes.rank == b.core.nodes.rank


@requires_numpy
def test_write_web_viz_html_writes_html_file(tmp_path: Path) -> None:
    g = _build_two_component_digraph()
    payload = to_web_viz_payload(g, seed=3)
    out = tmp_path / "web-viz.html"

    write_web_viz_html(payload, out, data_mode="inline")

    html = out.read_text(encoding="utf-8")
    assert "Directed Louvain modules" in html
    assert "webviz.louvain_directed" in html
    m = re.search(r"window\.__VIZ_DATA__\s*=\s*(\{.*?\});", html, re.S)
    assert m, "inline JSON"
    d = json.loads(m.group(1))
    assert d.get("annotations", {}).get("layout") == "clustered_force"
    assert "window.VIZ_LAYOUT_CONFIG = " in html
    assert "VizForce.relayoutGroup" in html
    assert "d3-force" not in html
    assert "40000" not in html


@requires_numpy
def test_write_web_viz_html_accepts_custom_template(tmp_path: Path) -> None:
    g = _build_two_component_digraph()
    p = to_web_viz_payload(g, seed=1)
    pkg = "excel_grapher.grapher"
    ref = importlib.resources.files(pkg).joinpath("lightweight_viz_template.html")
    tpl = tmp_path / "tpl.html"
    tpl.write_text(ref.read_text(encoding="utf-8"), encoding="utf-8")
    out = tmp_path / "out.html"
    write_web_viz_html(
        p,
        out,
        data_mode="inline",
        title="T",
        template_path=tpl,
    )
    assert out.is_file() and "T" in out.read_text(encoding="utf-8")
    assert "createREGL" in out.read_text(encoding="utf-8")


@pytest.mark.parametrize(
    "layout",
    ("spring", "forceatlas2"),
)
def test_to_web_viz_payload_supports_networkx_layouts(layout: str) -> None:
    # NetworkX spring/multipartite layouts import NumPy internally.
    pytest.importorskip("numpy")
    g = _build_two_component_digraph()
    payload = to_web_viz_payload(g, layout=layout, seed=11)
    assert payload.core.stats.node_count == 4
    assert any(
        abs(x) > 0 or abs(y) > 0
        for x, y in zip(payload.core.nodes.x, payload.core.nodes.y, strict=True)
    )


@requires_numpy
def test_clustered_force_positions_come_from_shared_layout() -> None:
    from excel_grapher.grapher.lightweight_viz import _build_int_adjacencies
    from excel_grapher.grapher.viz_layout import clustered_force_layout, input_depths

    graph = _build_two_component_graph()
    payload = to_web_viz_payload(graph, seed=5)
    keys = graph.keys(order="workbook")
    key_id = {k: i for i, k in enumerate(keys)}
    uncond, all_adj = _build_int_adjacencies(graph, keys, key_id)
    edges = [(u, v) for u in range(len(keys)) for v in all_adj[u]]
    depths = input_depths(len(keys), [(u, v) for u in range(len(keys)) for v in uncond[u]])
    modules = payload.overlays[0].data["node_module_id"]
    want = clustered_force_layout(
        len(keys), edges, modules, depths=depths, rank_pull="between", seed=5
    )
    assert list(payload.core.nodes.depth) == depths
    for i in range(len(keys)):
        assert payload.core.nodes.x[i] == pytest.approx(want[i][0])
        assert payload.core.nodes.y[i] == pytest.approx(want[i][1])


@requires_numpy
def test_clustered_force_rank_pull_is_configurable() -> None:
    graph = _build_chain_digraph(6)
    pulled = to_web_viz_payload(graph, seed=0)
    free = to_web_viz_payload(graph, seed=0, layout_config={"rank_pull": "none"})
    assert free.annotations is not None and free.annotations["rank_pull"] == "none"
    assert pulled.core.nodes.y != free.core.nodes.y
    with pytest.raises(ValueError, match="rank_pull"):
        to_web_viz_payload(graph, layout_config={"rank_pull": "sideways"})


@requires_numpy
def test_to_web_viz_payload_can_omit_module_overlay() -> None:
    g = _build_two_component_digraph()
    payload = to_web_viz_payload(g, include_module_overlay=False, seed=1)
    assert payload.overlays == ()
    assert payload.annotations is not None
    assert payload.annotations.get("layout") == "clustered_force"
    assert payload.core.stats.node_count == 4


def _graph_with_unreachable_cycle() -> DependencyGraph:
    """`S!D1` is the only BFS seed; the `A1 <-> B1` cycle is unreachable from it."""
    graph = DependencyGraph()
    graph.add_node(Node(sheet="S", column="D", row=1, formula=None, value=1, is_leaf=True))
    graph.add_node(Node(sheet="S", column="A", row=1, formula="=B1", value=None, is_leaf=False))
    graph.add_node(Node(sheet="S", column="B", row=1, formula="=A1", value=None, is_leaf=False))
    graph.add_edge("S!A1", "S!B1")
    graph.add_edge("S!B1", "S!A1")
    return graph


@pytest.mark.parametrize("include_module_overlay", [True, False])
@requires_numpy
def test_payload_node_columns_align_without_guarded_edges(include_module_overlay: bool) -> None:
    payload = to_web_viz_payload(
        _graph_with_unreachable_cycle(),
        include_guarded_edges=False,
        include_module_overlay=include_module_overlay,
    )
    nodes = payload.core.nodes
    n = payload.core.stats.node_count
    assert n == 3
    for column in (nodes.sheet_index, nodes.row, nodes.column, nodes.rank, nodes.x, nodes.y):
        assert len(column) == n
