"""NetworkX-first web viz API coverage (integration)."""

from __future__ import annotations

import importlib.resources
import inspect
import json
import re
from pathlib import Path

import pytest

from excel_grapher import write_web_viz_html
from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.exporter.web_viz_layout import LAYOUT_STRATIFIED_MULTIPARTITE, list_web_viz_layouts
from excel_grapher.grapher.lightweight_viz import lightweight_viz_flat


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


def test_to_web_viz_payload_default_layout_is_stratified_multipartite() -> None:
    sig = inspect.signature(to_web_viz_payload)
    assert sig.parameters["layout"].default == "stratified_multipartite"


def test_to_web_viz_layout_registry_includes_builtins() -> None:
    ids = set(list_web_viz_layouts())
    assert "stratified_multipartite" in ids
    assert "spring" in ids
    assert "multipartite" in ids


def test_to_web_viz_payload_unknown_layout_raises() -> None:
    g = _build_two_component_digraph()
    with pytest.raises(ValueError, match="Unknown web viz layout"):
        to_web_viz_payload(g, layout="not.a.registered.id", seed=0)


def test_to_web_viz_payload_includes_annotations() -> None:
    g = _build_two_component_digraph()
    payload = to_web_viz_payload(g, seed=7, layout=LAYOUT_STRATIFIED_MULTIPARTITE)
    assert payload.annotations is not None
    assert payload.annotations.get("layout") == "stratified_multipartite"

    from excel_grapher.grapher.lightweight_viz import serialize_lightweight_viz_json

    blob = json.loads(serialize_lightweight_viz_json(payload))
    assert blob.get("annotations", {}).get("layout") == "stratified_multipartite"


def test_to_web_viz_payload_accepts_networkx_digraph() -> None:
    g = _build_two_component_digraph()

    payload = to_web_viz_payload(g, seed=7)
    flat = lightweight_viz_flat(payload)

    assert flat.stats.node_count == 4
    assert flat.stats.module_count == 2
    assert payload.overlays[0].display_name == "Directed Louvain modules"
    assert payload.overlays[0].overlay_id == "webviz.louvain_directed"


def test_to_web_viz_payload_is_deterministic_with_seed() -> None:
    g = _build_two_component_digraph()

    a = lightweight_viz_flat(to_web_viz_payload(g, seed=17))
    b = lightweight_viz_flat(to_web_viz_payload(g, seed=17))

    assert list(a.nodes.module_id) == list(b.nodes.module_id)
    assert list(a.nodes.rank) == list(b.nodes.rank)


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
    assert d.get("annotations", {}).get("layout") == "stratified_multipartite"


def test_write_web_viz_html_accepts_custom_template(tmp_path: Path) -> None:
    g = _build_two_component_digraph()
    p = to_web_viz_payload(g, seed=1, layout="spring")
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
    ("spring", "forceatlas2", "multipartite"),
)
def test_to_web_viz_payload_supports_networkx_layouts(layout: str) -> None:
    g = _build_two_component_digraph()
    payload = to_web_viz_payload(g, layout=layout, seed=11)
    flat = lightweight_viz_flat(payload)
    assert flat.stats.node_count == 4
    assert any(abs(x) > 0 or abs(y) > 0 for x, y in zip(flat.nodes.x, flat.nodes.y, strict=True))


def test_to_web_viz_stratified_has_distinct_scc_ranks() -> None:
    g = _build_chain_digraph(5)
    flat = lightweight_viz_flat(
        to_web_viz_payload(g, seed=0, layout="stratified_multipartite", include_module_overlay=True)
    )
    assert flat.stats.node_count == 5
    ranks = set(flat.nodes.rank)
    assert len(ranks) >= 2


def test_to_web_viz_payload_can_omit_module_overlay() -> None:
    g = _build_two_component_digraph()
    payload = to_web_viz_payload(g, include_module_overlay=False, seed=1)
    assert payload.overlays == ()
    assert payload.annotations is not None
    assert payload.annotations.get("stratified_fallback") == "bfs_multipartite"
    flat = lightweight_viz_flat(payload)
    assert flat.stats.node_count == 4
    assert flat.stats.module_count == 1
