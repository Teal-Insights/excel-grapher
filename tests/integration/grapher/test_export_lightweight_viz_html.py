"""HTML web-viz output embeds coherent payloads from graph inputs (integration)."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

import excel_grapher.grapher.lightweight_viz as lightweight_viz_mod
from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.grapher import write_lightweight_viz_data, write_lightweight_viz_html
from excel_grapher.grapher.lightweight_viz import (
    VIZ_PAYLOAD_VERSION,
    VizLimits,
    assemble_lightweight_viz_payload,
    build_lightweight_viz_core,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node


def _n(sheet: str, col: str, row: int, *, leaf: bool, formula: str | None) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=1 if leaf else None,
        is_leaf=leaf,
    )


def _chain_graph() -> DependencyGraph:
    g = DependencyGraph()
    n1 = _n("S", "A", 1, leaf=True, formula=None)
    n2 = _n("S", "A", 2, leaf=False, formula="=A1")
    n3 = _n("S", "A", 3, leaf=False, formula="=A2")
    for n in (n1, n2, n3):
        g.add_node(n)
    g.add_edge(n3.key, n2.key)
    g.add_edge(n2.key, n1.key)
    return g


def _chain_nx():
    import networkx as nx

    g = nx.DiGraph()
    g.add_node("S!A1", formula=None, value=1, is_leaf=True, sheet="S", column="A", row=1)
    g.add_node("S!A2", formula="=A1", value=None, is_leaf=False, sheet="S", column="A", row=2)
    g.add_node("S!A3", formula="=A2", value=None, is_leaf=False, sheet="S", column="A", row=3)
    g.add_edge("S!A3", "S!A2")
    g.add_edge("S!A2", "S!A1")
    return g


def _payload():
    return to_web_viz_payload(_chain_nx(), layout="stratified_multipartite")


def test_write_html_core_only_no_overlays(tmp_path: Path) -> None:
    g = _chain_graph()
    core = build_lightweight_viz_core(g, limits=VizLimits(), layout_input=None)
    payload = assemble_lightweight_viz_payload(core, [])
    assert payload.version == VIZ_PAYLOAD_VERSION
    assert payload.overlays == ()
    out = tmp_path / "core_only.html"
    write_lightweight_viz_html(payload, out, title="Core only", data_mode="inline")
    assert out.is_file()
    text = out.read_text(encoding="utf-8")
    ver = str(VIZ_PAYLOAD_VERSION)
    assert f'"version":{ver}' in text or f'"version": {ver}' in text.replace(" ", "")
    assert "Core only" in text


def test_write_html_creates_file(tmp_path: Path) -> None:
    p = _payload()
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, title="T", data_mode="inline")
    assert out.is_file()
    text = out.read_text(encoding="utf-8")
    assert "T" in text
    assert "canvas" in text
    assert "createREGL" in text or "regl" in text.lower()
    assert "d3.forceSimulation" in text or "d3-force" in text


def test_inline_embeds_payload_under_budget(tmp_path: Path) -> None:
    p = _payload()
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, data_mode="inline", inline_size_budget_mb=50)
    text = out.read_text(encoding="utf-8")
    assert "window.__VIZ_DATA__" in text
    ver = str(VIZ_PAYLOAD_VERSION)
    assert f'"version":{ver}' in text or f'"version": {ver}' in text.replace(" ", "")
    assert '"formula"' in text


def test_sidecar_writes_sibling_json(tmp_path: Path) -> None:
    p = _payload()
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, data_mode="sidecar", data_path=tmp_path / "data.viz.json")
    data = tmp_path / "data.viz.json"
    assert data.is_file()
    assert "window.__VIZ_DATA_URL__" in out.read_text(encoding="utf-8")


def test_auto_sidecar_when_estimate_large(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    p = _payload()
    monkeypatch.setattr(
        lightweight_viz_mod,
        "estimate_serialized_json_bytes",
        lambda _payload: 100 * 1024 * 1024,
    )
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, data_mode="auto", inline_size_budget_mb=50)
    sidecar = tmp_path / "v.viz.json"
    assert sidecar.is_file()
    html = out.read_text(encoding="utf-8")
    assert "__VIZ_DATA_URL__" in html


def test_invalid_payload_version_raises(tmp_path: Path) -> None:
    p = replace(_payload(), version=99)
    with pytest.raises(ValueError, match="Unsupported"):
        write_lightweight_viz_html(p, tmp_path / "x.html", data_mode="inline")


def test_write_data_roundtrip(tmp_path: Path) -> None:
    p = _payload()
    path = tmp_path / "d.json"
    write_lightweight_viz_data(p, path)
    assert path.read_text(encoding="utf-8").startswith("{")


@pytest.mark.parametrize(
    "needle",
    [
        "regl",
        "d3-force",
        "module_edges",
        "Local force",
    ],
)
def test_overview_viewer_contract(tmp_path: Path, needle: str) -> None:
    p = _payload()
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, data_mode="inline")
    assert needle.lower() in out.read_text(encoding="utf-8").lower()
