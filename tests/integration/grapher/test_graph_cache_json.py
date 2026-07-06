"""Graph cache JSON round-trips stay aligned with freshly built graphs (integration).

Covers `build_graph_cache_meta`, `save_graph_cache`, and `try_load_graph_cache`
on disk-backed workbooks so cache hits, invalidation, and validation policy behave
predictably across releases.
"""

from __future__ import annotations

import json
from pathlib import Path

import xlsxwriter

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.grapher.cache import (
    GRAPH_CACHE_SCHEMA_VERSION,
    CacheValidationPolicy,
    build_graph_cache_meta,
    build_graph_cache_meta_portable,
    save_graph_cache,
    try_load_graph_cache,
)


def _make_simple_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 2)  # A1
    ws.write_number(1, 0, 3)  # A2
    ws.write_formula(2, 0, "=A1+A2", None, 5)  # A3
    wb.close()


def test_cache_roundtrip_date_and_time_leaf_values(tmp_path: Path) -> None:
    """Excel cached values may include date/time objects; graph cache must round-trip them."""
    from datetime import date, time

    from excel_grapher.grapher.cache import _value_from_json, _value_to_json
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.node import Node

    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(
        Node(
            sheet="Sheet1",
            column="A",
            row=1,
            formula=None,
            normalized_formula=None,
            value=date(2025, 8, 12),
            is_leaf=True,
            is_target=False,
        )
    )
    graph.add_node(
        Node(
            sheet="Sheet1",
            column="B",
            row=1,
            formula=None,
            normalized_formula=None,
            value=time(9, 30),
            is_leaf=True,
            is_target=False,
        )
    )

    workbook = tmp_path / "book.xlsx"
    _make_simple_workbook(workbook)
    meta = build_graph_cache_meta(workbook, ["Sheet1!A1"], extraction_params={})
    cache_path = tmp_path / "typed-leaves.json"
    save_graph_cache(cache_path, graph, meta)

    loaded = try_load_graph_cache(cache_path, expected_meta=meta)
    assert loaded is not None
    node_a1 = loaded.get_node("Sheet1!A1")
    node_b1 = loaded.get_node("Sheet1!B1")
    assert node_a1 is not None
    assert node_b1 is not None
    assert node_a1.value == date(2025, 8, 12)
    assert node_b1.value == time(9, 30)

    assert _value_from_json(_value_to_json(date(2024, 1, 2))) == date(2024, 1, 2)
    assert _value_from_json(_value_to_json(time(12, 0))) == time(12, 0)


def test_cache_roundtrip_strict_hit(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    _make_simple_workbook(workbook)

    targets = ["Sheet1!A3"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    meta = build_graph_cache_meta(
        workbook,
        targets,
        extraction_params={"load_values": True, "max_depth": None},
    )

    cache_path = tmp_path / "graph.json"
    save_graph_cache(cache_path, graph, meta)

    loaded = try_load_graph_cache(cache_path, expected_meta=meta)
    assert loaded is not None
    assert loaded.get_dependencies("Sheet1!A3") == {"Sheet1!A1", "Sheet1!A2"}

    with FormulaEvaluator(loaded) as ev:
        assert ev.evaluate(["Sheet1!A3"])["Sheet1!A3"] == 5


def test_cache_miss_when_targets_differ(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    _make_simple_workbook(workbook)

    targets = ["Sheet1!A3"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    meta = build_graph_cache_meta(workbook, targets, extraction_params={})

    cache_path = tmp_path / "graph.json"
    save_graph_cache(cache_path, graph, meta)

    meta_other_targets = build_graph_cache_meta(workbook, ["Sheet1!A2"], extraction_params={})
    loaded = try_load_graph_cache(cache_path, expected_meta=meta_other_targets)
    assert loaded is None


def test_cache_invalidation_when_extraction_params_differ(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    _make_simple_workbook(workbook)

    targets = ["Sheet1!A3"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    meta = build_graph_cache_meta(workbook, targets, extraction_params={"max_depth": 50})

    cache_path = tmp_path / "graph.json"
    save_graph_cache(cache_path, graph, meta)

    meta_other_params = build_graph_cache_meta(
        workbook, targets, extraction_params={"max_depth": 51}
    )
    loaded = try_load_graph_cache(cache_path, expected_meta=meta_other_params)
    assert loaded is None


def test_cache_schema_version_mismatch(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    _make_simple_workbook(workbook)

    targets = ["Sheet1!A3"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    meta = build_graph_cache_meta(workbook, targets, extraction_params={})

    cache_path = tmp_path / "graph.json"
    save_graph_cache(cache_path, graph, meta)

    payload = json.loads(cache_path.read_text(encoding="utf-8"))
    payload["meta"]["schema_version"] = GRAPH_CACHE_SCHEMA_VERSION - 1
    cache_path.write_text(json.dumps(payload), encoding="utf-8")

    loaded = try_load_graph_cache(cache_path, expected_meta=meta)
    assert loaded is None


def test_portable_load_skips_workbook_checks_but_enforces_others(tmp_path: Path) -> None:
    """Enforce portable graph-cache validation without the workbook file.

    Graph load must not require access to the workbook file, but it must still
    enforce schema, excel-grapher version, targets, and extraction params.
    """
    workbook = tmp_path / "book.xlsx"
    _make_simple_workbook(workbook)

    targets = ["Sheet1!A3"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    meta = build_graph_cache_meta(workbook, targets, extraction_params={"load_values": True})

    cache_path = tmp_path / "graph.json"
    save_graph_cache(cache_path, graph, meta)

    portable_expected = build_graph_cache_meta_portable(
        targets,
        extraction_params={"load_values": True},
    )

    loaded = try_load_graph_cache(
        cache_path,
        expected_meta=portable_expected,
        policy=CacheValidationPolicy.PORTABLE,
    )
    assert loaded is not None
    with FormulaEvaluator(loaded) as ev:
        assert ev.evaluate(["Sheet1!A3"])["Sheet1!A3"] == 5

    # Still rejects if targets differ.
    portable_other_targets = build_graph_cache_meta_portable(
        ["Sheet1!A2"],
        extraction_params={"load_values": True},
    )
    assert (
        try_load_graph_cache(
            cache_path,
            expected_meta=portable_other_targets,
            policy=CacheValidationPolicy.PORTABLE,
        )
        is None
    )

    # Still rejects if extraction params differ.
    portable_other_params = build_graph_cache_meta_portable(
        targets,
        extraction_params={"load_values": False},
    )
    assert (
        try_load_graph_cache(
            cache_path,
            expected_meta=portable_other_params,
            policy=CacheValidationPolicy.PORTABLE,
        )
        is None
    )
