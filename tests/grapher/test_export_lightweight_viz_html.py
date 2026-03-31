from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

import excel_grapher.grapher.lightweight_viz as lightweight_viz_mod
from excel_grapher.grapher import (
    to_lightweight_viz,
    write_lightweight_viz_data,
    write_lightweight_viz_html,
)
from tests.grapher.test_export_lightweight_viz import _chain_graph


def test_write_html_creates_file(tmp_path: Path) -> None:
    p = to_lightweight_viz(_chain_graph())
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, title="T", data_mode="inline")
    assert out.is_file()
    text = out.read_text(encoding="utf-8")
    assert "T" in text
    assert "canvas" in text
    assert "createREGL" in text or "regl" in text.lower()
    assert "d3.forceSimulation" in text or "d3-force" in text


def test_inline_embeds_payload_under_budget(tmp_path: Path) -> None:
    p = to_lightweight_viz(_chain_graph())
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, data_mode="inline", inline_size_budget_mb=50)
    text = out.read_text(encoding="utf-8")
    assert "window.__VIZ_DATA__" in text
    assert '"version":1' in text or '"version": 1' in text.replace(" ", "")


def test_sidecar_writes_sibling_json(tmp_path: Path) -> None:
    p = to_lightweight_viz(_chain_graph())
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, data_mode="sidecar", data_path=tmp_path / "data.viz.json")
    data = tmp_path / "data.viz.json"
    assert data.is_file()
    assert "window.__VIZ_DATA_URL__" in out.read_text(encoding="utf-8")


def test_auto_sidecar_when_estimate_large(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    p = to_lightweight_viz(_chain_graph())
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
    p = replace(to_lightweight_viz(_chain_graph()), version=99)
    with pytest.raises(ValueError, match="Unsupported"):
        write_lightweight_viz_html(p, tmp_path / "x.html", data_mode="inline")


def test_write_data_roundtrip(tmp_path: Path) -> None:
    p = to_lightweight_viz(_chain_graph())
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
    p = to_lightweight_viz(_chain_graph())
    out = tmp_path / "v.html"
    write_lightweight_viz_html(p, out, data_mode="inline")
    assert needle.lower() in out.read_text(encoding="utf-8").lower()
