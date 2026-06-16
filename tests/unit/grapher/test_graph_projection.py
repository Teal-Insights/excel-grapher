"""Unit tests for graph projection and identity transit compression."""

from __future__ import annotations

from pathlib import Path
from typing import cast

import pytest
import xlsxwriter

from excel_grapher.exporter.projection import (
    IdentityTransitCompression,
    ProjectionManifest,
    ProjectionResult,
    apply_projection,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.node import Node
from excel_grapher.series_bindings import WorkbookSeriesBindings, resolve_series_bindings


def _make_node(
    key: str,
    formula: str | None,
    normalized: str | None,
    *,
    is_leaf: bool = False,
    is_target: bool = False,
) -> Node:
    sheet, rest = key.split("!", 1)
    if sheet.startswith("'"):
        sheet = sheet[1:-1]
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


def test_identity_projection_does_not_mutate_original_graph() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 42)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    original_keys = set(graph)
    projection = IdentityTransitCompression().project(graph)

    assert set(graph) == original_keys
    assert "Sheet1!B1" in graph
    assert "Sheet1!B1" not in projection
    assert projection.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!C1"})


def test_projection_result_behaves_like_projected_graph() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    projection = IdentityTransitCompression().project(graph)

    assert "Sheet1!B1" not in projection
    assert projection.get_node("Sheet1!C1") is not None
    assert projection.leaf_keys() == ["Sheet1!C1"]
    assert projection.original_graph is graph
    assert projection.projected_graph is not projection
    assert "Sheet1!B1" in projection.original_graph


def test_manifest_serializes_and_resolves_chain_aliases() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    c = _make_node("Sheet1!C1", "=Sheet1!D1", "=Sheet1!D1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (d, c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!C1", "Sheet1!D1", provenance=EdgeProvenance(causes=frozenset({dr})))
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    manifest = IdentityTransitCompression().project(graph).manifest

    assert manifest.removed_to_replacement["Sheet1!B1"] == "Sheet1!D1"
    assert manifest.removed_to_replacement["Sheet1!C1"] == "Sheet1!D1"
    assert manifest.retained_to_collapsed_sources["Sheet1!D1"] == ("Sheet1!B1", "Sheet1!C1")

    restored = type(manifest).from_dict(manifest.to_dict())
    assert restored.removed_to_replacement == manifest.removed_to_replacement
    assert restored.collapsed_groups == manifest.collapsed_groups


def test_collapsed_group_records_statement_order_and_external_boundary() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    c = _make_node("Sheet1!C1", "=Sheet1!D1", "=Sheet1!D1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (d, c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!C1", "Sheet1!D1", provenance=EdgeProvenance(causes=frozenset({dr})))
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    groups = IdentityTransitCompression().project(graph).manifest.collapsed_groups
    assert len(groups) >= 1
    group = groups[-1]
    assert group.retained == "Sheet1!D1"
    assert group.collapsed_sources == ("Sheet1!B1", "Sheet1!C1")
    assert group.statement_order == ("Sheet1!C1", "Sheet1!B1")
    assert "Sheet1!D1" not in group.external_dependencies
    assert any(snapshot.address in {"Sheet1!B1", "Sheet1!C1"} for snapshot in group.node_snapshots)


def test_apply_projection_preserves_manifest_from_earlier_steps() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    c = _make_node("Sheet1!C1", "=Sheet1!D1", "=Sheet1!D1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (d, c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!C1", "Sheet1!D1", provenance=EdgeProvenance(causes=frozenset({dr})))
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    projection = apply_projection(
        graph,
        [IdentityTransitCompression(), IdentityTransitCompression()],
    )

    assert projection.manifest.removed_to_replacement["Sheet1!B1"] == "Sheet1!D1"
    assert projection.manifest.removed_to_replacement["Sheet1!C1"] == "Sheet1!D1"


@pytest.mark.xfail(
    reason="apply_projection only composes identity-transit manifests today",
    strict=True,
)
def test_apply_projection_rejects_heterogeneous_projection_manifests() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    class HeterogeneousProjection:
        def project(self, graph: DependencyGraph) -> ProjectionResult:
            manifest = ProjectionManifest.empty()
            object.__setattr__(manifest, "kind", "subgraph")
            return ProjectionResult(
                original_graph=graph,
                projected_graph=graph,
                manifest=manifest,
            )

    graph = DependencyGraph()
    graph.add_node(_make_node("Sheet1!A1", None, None, is_leaf=True))

    with pytest.raises(NotImplementedError, match="heterogeneous projection"):
        apply_projection(graph, [HeterogeneousProjection()])


def test_manifest_node_snapshots_preserve_cell_coordinates_and_values() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 42)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    object.__setattr__(b, "value", 42)
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    manifest_dict = IdentityTransitCompression().project(graph).manifest.to_dict()
    snapshot = manifest_dict["collapsed_groups"][0]["node_snapshots"][0]

    assert snapshot["sheet"] == "Sheet1"
    assert snapshot["column"] == "B"
    assert snapshot["row"] == 1
    assert snapshot["value"] == 42


def test_issue_224_mcve_bindings_resolve_on_original_graph(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    wb = xlsxwriter.Workbook(workbook_path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("C6", 10)
    out = wb.add_worksheet("Outputs")
    out.write_formula("B12", "=Engine!C6")
    out.write_formula("B14", "=Outputs!B12+1")
    wb.close()

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B12", "Outputs!B14"],
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
                }
            ],
        },
    )

    before = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook_path,
        direction="output",
    )
    assert before["series"][0]["leaves"]

    projection = IdentityTransitCompression().project(graph)
    assert "Outputs!B12" not in projection
    assert "Outputs!B12" in graph

    after = resolve_series_bindings(
        projection.original_graph,
        bindings,
        workbook=workbook_path,
        direction="output",
    )
    assert after["series"][0]["leaves"]


def test_static_range_blocks_projection(tmp_path: Path) -> None:
    path = tmp_path / "rng.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 1)
    ws.write_number(0, 2, 2)
    ws.write_formula(0, 0, "=SUM(Sheet1!B1:C1)", None, 3)
    ws.write_formula(0, 3, "=Sheet1!B1", None, 1)
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Sheet1!A1"],
        load_values=False,
        capture_dependency_provenance=True,
    )
    projection = IdentityTransitCompression().project(graph)
    assert "Sheet1!B1" in projection
    assert not projection.manifest.removed_to_replacement


@pytest.mark.parametrize(
    "test_name",
    ["guard", "provenance_absent"],
)
def test_projection_respects_compression_safety(test_name: str) -> None:
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.guard import Literal

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    if test_name == "guard":
        graph.add_edge(
            "Sheet1!B1",
            "Sheet1!C1",
            guard=Literal(True),
            provenance=EdgeProvenance(causes=frozenset({dr})),
        )
        graph.add_edge(
            "Sheet1!A1",
            "Sheet1!B1",
            provenance=EdgeProvenance(
                causes=frozenset({dr}),
                direct_sites_formula=sp,
                direct_sites_normalized=sp,
            ),
        )
    else:
        graph.add_edge("Sheet1!B1", "Sheet1!C1")
        graph.add_edge("Sheet1!A1", "Sheet1!B1")

    projection = IdentityTransitCompression().project(graph)
    assert "Sheet1!B1" in projection
    assert not projection.manifest.removed_to_replacement
