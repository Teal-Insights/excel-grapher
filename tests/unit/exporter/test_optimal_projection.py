"""Unit tests for optimal compression projection."""

from __future__ import annotations

from pathlib import Path
from typing import cast

import pytest
import xlsxwriter

from excel_grapher.core.formula_ast import parse
from excel_grapher.exporter.projection import (
    BaseProjectionManifest,
    OptimalCompression,
    _statement_order,
    _statement_order_index,
    build_forwarding_projection_manifest,
    build_optimal_projection_manifest,
    resolve_projection_manifest,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.compression import (
    IdentityTransitCompressionRecord,
    OptimalCompressionRecord,
)
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef, Compare, Literal
from excel_grapher.grapher.node import Node
from excel_grapher.series_bindings.types import WorkbookSeriesBindings


def _make_node(
    key: str,
    formula: str | None,
    normalized: str | None,
    *,
    is_leaf: bool = False,
    is_target: bool = False,
) -> Node:
    sheet, rest = key.split("!", 1)
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=normalized,
        value=None,
        is_leaf=is_leaf,
        is_target=is_target,
    )


def _direct_edge(graph: DependencyGraph, dependent: str, precedent: str) -> None:
    dr = DependencyCause.direct_ref
    dep_node = graph.get_node(dependent)
    assert dep_node is not None
    n = dep_node.normalized_formula
    assert n is not None
    ref = precedent
    i_n = n.index(ref)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=dr,
            direct_sites_normalized=((i_n, i_n + len(ref)),),
        ),
    )


def test_optimal_projection_does_not_mutate_original_graph() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")
    graph.set_node_metadata("Sheet1!B1", {"label": "source"})
    ast = parse("=Sheet1!D1*2")
    graph.preparsed_formulas = {"=Sheet1!D1*2": ast}

    original_keys = set(graph)
    projection = OptimalCompression().project(graph)

    assert set(graph) == original_keys
    assert "Sheet1!B1" in graph
    assert "Sheet1!B1" not in projection
    original_a = graph.get_node("Sheet1!A1")
    original_b = graph.get_node("Sheet1!B1")
    projected_a = projection.get_node("Sheet1!A1")
    assert original_a is not None
    assert original_b is not None
    assert projected_a is not None
    assert original_a.normalized_formula == "=Sheet1!B1+1"
    assert original_b.normalized_formula == "=Sheet1!D1*2"
    assert original_b.metadata["label"] == "source"
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})
    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!D1"})
    assert graph.preparsed_formulas == {"=Sheet1!D1*2": ast}
    assert projected_a.normalized_formula == "=(Sheet1!D1*2)+1"
    assert projection.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!D1"})


def test_optimal_projection_manifest_kind_and_forwarding() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    manifest = OptimalCompression().project(graph).manifest
    projection = OptimalCompression().project(graph)
    assert isinstance(manifest, BaseProjectionManifest)
    assert manifest.kind == "optimal_compression"
    assert manifest.map_to_projected("Sheet1!B1") == "Sheet1!C1"
    assert "Sheet1!C1" in projection


def test_optimal_projection_keeps_target_identity_transit() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1", is_target=True)
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    projection = OptimalCompression().project(graph)
    assert "Sheet1!B1" in projection
    assert projection.manifest.map_to_projected("Sheet1!B1") == "Sheet1!B1"


def test_optimal_projection_preserves_series_bound_non_target(
    tmp_path: Path,
) -> None:
    workbook_path = tmp_path / "series_bound.xlsx"
    wb = xlsxwriter.Workbook(workbook_path)
    engine = wb.add_worksheet("Engine")
    engine.write_number("C6", 10)
    out = wb.add_worksheet("Outputs")
    out.write_formula("B12", "=Engine!C6", None, 10)
    out.write_formula("B14", "=Outputs!B12+1", None, 11)
    wb.close()

    # Only B14 is an export target; B12 is published solely via series bindings.
    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B14"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    bindings = cast(
        WorkbookSeriesBindings,
        {
            "schema_version": "1.2.0",
            "workbook": str(workbook_path),
            "series": [
                {
                    "id": "baseline",
                    "data_range": "Outputs!B12",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_baseline"}},
                    "structure": {
                        "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                        "dimensions": [
                            {
                                "concept": "LABEL",
                                "role": "key",
                                "scope": "series",
                                "bind": {"kind": "constant", "value": "baseline"},
                            }
                        ],
                    },
                    "key": ["LABEL"],
                }
            ],
        },
    )

    without_preserve = OptimalCompression().project(graph)
    assert "Outputs!B12" not in without_preserve

    projection = OptimalCompression(
        series_bindings=bindings,
        bindings_workbook=workbook_path,
    ).project(graph)
    assert "Outputs!B12" in projection
    assert projection.manifest.map_to_projected("Outputs!B12") == "Outputs!B12"


def test_optimal_projection_inline_lineage_without_forwarding() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    manifest = OptimalCompression().project(graph).manifest
    assert isinstance(manifest, BaseProjectionManifest)
    assert manifest.map_to_projected("Sheet1!B1") == "Sheet1!B1"
    assert manifest.retained_to_collapsed_sources["Sheet1!A1"] == ("Sheet1!B1",)
    assert "Sheet1!B1" in manifest.removed_node_snapshots


def test_optimal_projection_resolves_chained_inline_to_final_retained() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    c = _make_node("Sheet1!C1", "=Sheet1!D1+1", "=Sheet1!D1+1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1*2", "=Sheet1!C1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+3", "=Sheet1!B1+3")
    for n in (d, c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!C1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    manifest = OptimalCompression().project(graph).manifest
    assert isinstance(manifest, BaseProjectionManifest)
    collapsed = set(manifest.retained_to_collapsed_sources["Sheet1!A1"])
    assert collapsed == {"Sheet1!B1", "Sheet1!C1"}


def test_optimal_manifest_reuses_statement_order_across_groups(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = DependencyGraph()
    record = OptimalCompressionRecord()
    group_count = 8
    for row in range(1, group_count + 1):
        retained = f"Sheet1!A{row}"
        source = f"Sheet1!B{row}"
        leaf = f"Sheet1!C{row}"
        graph.add_node(_make_node(leaf, None, None, is_leaf=True))
        graph.add_node(_make_node(source, f"={leaf}+1", f"={leaf}+1"))
        graph.add_node(_make_node(retained, f"={source}+1", f"={source}+1"))
        _direct_edge(graph, source, leaf)
        _direct_edge(graph, retained, source)
        record.inlined_to[source] = retained

    calls = 0
    original_evaluation_order = graph.evaluation_order

    def counted_evaluation_order(
        *, strict: bool = True, iterate_enabled: bool | None = None
    ) -> list[str]:
        nonlocal calls
        calls += 1
        return original_evaluation_order(strict=strict, iterate_enabled=iterate_enabled)

    monkeypatch.setattr(graph, "evaluation_order", counted_evaluation_order)

    manifest = build_optimal_projection_manifest(graph, record)

    assert len(manifest.collapsed_groups) == group_count
    assert calls == 1


def test_forwarding_manifest_reuses_statement_order_across_groups(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = DependencyGraph()
    record = IdentityTransitCompressionRecord()
    group_count = 8
    for row in range(1, group_count + 1):
        dependent = f"Sheet1!A{row}"
        source = f"Sheet1!B{row}"
        retained = f"Sheet1!C{row}"
        graph.add_node(_make_node(retained, None, None, is_leaf=True))
        graph.add_node(_make_node(source, f"={retained}", f"={retained}"))
        graph.add_node(_make_node(dependent, f"={source}", f"={source}"))
        _direct_edge(graph, source, retained)
        _direct_edge(graph, dependent, source)
        record.immediate_removed[source] = retained
        record.removal_order.append(source)

    calls = 0
    original_evaluation_order = graph.evaluation_order

    def counted_evaluation_order(
        *, strict: bool = True, iterate_enabled: bool | None = None
    ) -> list[str]:
        nonlocal calls
        calls += 1
        return original_evaluation_order(strict=strict, iterate_enabled=iterate_enabled)

    monkeypatch.setattr(graph, "evaluation_order", counted_evaluation_order)

    manifest = build_forwarding_projection_manifest(graph, record, kind="identity_transit")

    assert len(manifest.collapsed_groups) == group_count
    assert calls == 1


def test_statement_order_cached_matches_uncached_may_cycle_fallback() -> None:
    graph = DependencyGraph()
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    b = _make_node("Sheet1!B1", "=Sheet1!A1+1", "=Sheet1!A1+1")
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    for n in (a, b, c):
        graph.add_node(n)
    provenance = EdgeProvenance(causes=DependencyCause.direct_ref)
    guard = Compare(left=CellRef(key="Sheet1!C1"), op="=", right=Literal(value=True))
    graph.add_edge("Sheet1!A1", "Sheet1!B1", guard=guard, provenance=provenance)
    graph.add_edge("Sheet1!B1", "Sheet1!A1", guard=guard, provenance=provenance)

    with pytest.warns(UserWarning, match="May-cycles detected"):
        uncached = _statement_order(graph, ("Sheet1!B1", "Sheet1!A1"))
    with pytest.warns(UserWarning, match="May-cycles detected"):
        order_index = _statement_order_index(graph)
    cached = _statement_order(
        graph,
        ("Sheet1!B1", "Sheet1!A1"),
        order_index=order_index,
    )

    assert cached == uncached == ("Sheet1!A1", "Sheet1!B1")


def test_optimal_projection_manifest_round_trips() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    manifest = OptimalCompression().project(graph).manifest
    assert isinstance(manifest, BaseProjectionManifest)
    restored = resolve_projection_manifest(manifest.to_dict())
    assert isinstance(restored, BaseProjectionManifest)
    assert restored.kind == "optimal_compression"
    assert restored.forwarding_map == manifest.forwarding_map
    assert restored.retained_to_collapsed_sources == manifest.retained_to_collapsed_sources
