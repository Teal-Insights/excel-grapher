"""Integration tests for optional label detection on ``create_dependency_graph``."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import (
    LabelDetectionConfig,
    create_dependency_graph,
    label_detection_config_to_jsonable,
)
from excel_grapher.grapher.cache import (
    build_graph_cache_meta,
    save_graph_cache,
    try_load_graph_cache,
)


def test_create_dependency_graph_with_label_detection_heuristic(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_string(0, 0, "Revenue")
    ws.write_number(0, 1, 2024)
    ws.write_formula(0, 2, "=B1*2", None, 0)
    wb.close()

    cfg = LabelDetectionConfig(enabled=True)
    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=True, label_detection=cfg)
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert node.metadata["row_labels"] == ["2024", "Revenue"]
    assert node.metadata["column_labels"] == []


def test_create_dependency_graph_without_label_detection_empty_metadata(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_string(0, 0, "Revenue")
    ws.write_number(0, 1, 2024)
    ws.write_formula(0, 2, "=B1*2", None, 0)
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=True)
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert dict(node.metadata) == {}


def test_graph_cache_roundtrip_with_label_detection_extraction_params(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_string(0, 0, "Revenue")
    ws.write_number(0, 1, 2024)
    ws.write_formula(0, 2, "=B1*2", None, 0)
    wb.close()

    targets = ["Sheet1!C1"]
    ld_cfg = LabelDetectionConfig(enabled=True)
    extraction_params = {
        "load_values": True,
        "label_detection": label_detection_config_to_jsonable(ld_cfg),
    }
    graph = create_dependency_graph(path, targets, load_values=True, label_detection=ld_cfg)
    meta = build_graph_cache_meta(path, list(targets), extraction_params=extraction_params)
    cache_path = tmp_path / "graph.json"
    save_graph_cache(cache_path, graph, meta)

    loaded = try_load_graph_cache(cache_path, expected_meta=meta)
    assert loaded is not None
    n = loaded.get_node("Sheet1!C1")
    assert n is not None
    assert n.metadata["row_labels"] == ["2024", "Revenue"]


def test_graph_cache_miss_when_label_detection_params_differ(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(0, 0, "=1+1", None, 2)
    wb.close()

    targets = ["Sheet1!A1"]
    ld_on = LabelDetectionConfig(enabled=True)
    ld_off = LabelDetectionConfig(enabled=False)
    params_on = {"label_detection": label_detection_config_to_jsonable(ld_on)}
    params_off = {"label_detection": label_detection_config_to_jsonable(ld_off)}

    graph = create_dependency_graph(path, targets, load_values=True, label_detection=ld_on)
    meta = build_graph_cache_meta(path, list(targets), extraction_params=params_on)
    cache_path = tmp_path / "graph.json"
    save_graph_cache(cache_path, graph, meta)

    meta_other = build_graph_cache_meta(path, list(targets), extraction_params=params_off)
    assert try_load_graph_cache(cache_path, expected_meta=meta_other) is None
