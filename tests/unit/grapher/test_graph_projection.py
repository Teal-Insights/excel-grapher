"""Unit tests for graph projection and identity transit compression."""

from __future__ import annotations

from dataclasses import fields
from pathlib import Path
from typing import Any, cast

import pytest
import xlsxwriter

from excel_grapher.core.formula_ast import FunctionCallNode, parse
from excel_grapher.exporter.projection import (
    BaseProjectionManifest,
    CompositeProjectionManifest,
    IdentityTransitCompression,
    OptimalCompression,
    ProjectedNodeSnapshot,
    ProjectionResult,
    apply_projection,
    register_projection_manifest,
    resolve_projection_manifest,
    unregister_projection_manifest,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.compression import CompressionProvenanceRequiredError
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


def test_resolve_projection_manifest_rejects_unknown_kind() -> None:
    with pytest.raises(ValueError, match="unsupported projection manifest kind"):
        resolve_projection_manifest({"kind": "subgraph"})


def test_base_projection_manifest_requires_explicit_empty_kind() -> None:
    with pytest.raises(TypeError):
        cast(Any, BaseProjectionManifest.empty)()


def test_register_and_resolve_custom_projection_manifest_round_trips() -> None:
    register_projection_manifest("custom_collapse", BaseProjectionManifest.from_dict)
    try:
        manifest = BaseProjectionManifest.empty(kind="custom_collapse")
        manifest.forwarding_map["Sheet1!B1"] = "Sheet1!C1"
        restored = resolve_projection_manifest(manifest.to_dict())
        assert isinstance(restored, BaseProjectionManifest)
        assert restored.kind == "custom_collapse"
        assert restored.forwarding_map == {"Sheet1!B1": "Sheet1!C1"}
    finally:
        unregister_projection_manifest("custom_collapse")


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
    assert isinstance(manifest, BaseProjectionManifest)

    assert manifest.forwarding_map["Sheet1!B1"] == "Sheet1!D1"
    assert manifest.forwarding_map["Sheet1!C1"] == "Sheet1!D1"
    assert manifest.retained_to_collapsed_sources["Sheet1!D1"] == ("Sheet1!B1", "Sheet1!C1")

    restored = resolve_projection_manifest(manifest.to_dict())
    assert isinstance(restored, BaseProjectionManifest)
    assert restored.forwarding_map == manifest.forwarding_map
    assert restored.collapsed_groups == manifest.collapsed_groups
    assert restored.removed_node_snapshots.keys() == manifest.removed_node_snapshots.keys()


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

    manifest = IdentityTransitCompression().project(graph).manifest
    assert isinstance(manifest, BaseProjectionManifest)
    groups = manifest.collapsed_groups
    assert len(groups) >= 1
    group = groups[-1]
    assert group.retained == "Sheet1!D1"
    assert group.collapsed_sources == ("Sheet1!B1", "Sheet1!C1")
    assert group.statement_order == ("Sheet1!C1", "Sheet1!B1")
    assert "Sheet1!D1" not in group.external_dependencies
    assert {"Sheet1!B1", "Sheet1!C1"} <= manifest.removed_node_snapshots.keys()


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

    assert projection.manifest.map_to_projected("Sheet1!B1") == "Sheet1!D1"
    assert projection.manifest.map_to_projected("Sheet1!C1") == "Sheet1!D1"


def test_apply_projection_composes_heterogeneous_kinds() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    class TagProjection:
        """Custom-kind step that records a tag without removing nodes."""

        def project(self, graph: DependencyGraph) -> ProjectionResult:
            return ProjectionResult(
                original_graph=graph,
                projected_graph=graph.copy(),
                manifest=BaseProjectionManifest.empty(kind="custom_tag"),
            )

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

    projection = apply_projection(graph, [IdentityTransitCompression(), TagProjection()])

    assert isinstance(projection.manifest, CompositeProjectionManifest)
    assert projection.manifest.map_to_projected("Sheet1!B1") == "Sheet1!C1"
    assert [step.kind for step in projection.manifest.steps] == ["identity_transit", "custom_tag"]

    register_projection_manifest("custom_tag", BaseProjectionManifest.from_dict)
    try:
        restored = resolve_projection_manifest(projection.manifest.to_dict())
    finally:
        unregister_projection_manifest("custom_tag")
    assert isinstance(restored, CompositeProjectionManifest)
    assert restored.forwarding_map == projection.manifest.forwarding_map
    assert [step.kind for step in restored.steps] == ["identity_transit", "custom_tag"]


def test_composite_projection_uses_manifest_protocol_for_mapping() -> None:
    class MapOnlyManifest:
        kind = "map_only"

        def map_to_projected(self, address: str) -> str:
            return "Sheet1!C1" if address == "Sheet1!B1" else address

        def to_dict(self) -> dict[str, object]:
            return {"kind": self.kind}

    manifest = CompositeProjectionManifest(forwarding_map={}, steps=(MapOnlyManifest(),))

    assert manifest.map_to_projected("Sheet1!B1") == "Sheet1!C1"
    assert manifest.forwarding_map == {}


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
    snapshot = manifest_dict["removed_node_snapshots"]["Sheet1!B1"]

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
    assert projection.manifest.map_to_projected("Sheet1!B1") == "Sheet1!B1"


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
        with pytest.raises(CompressionProvenanceRequiredError, match="provenance"):
            IdentityTransitCompression().project(graph)
        return

    projection = IdentityTransitCompression().project(graph)
    assert "Sheet1!B1" in projection
    assert projection.manifest.map_to_projected("Sheet1!B1") == "Sheet1!B1"


def test_custom_collapse_projection_uses_public_primitives_without_forwarding() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 5)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1", is_target=True)
    for n in (d, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!B1", "Sheet1!D1", provenance=EdgeProvenance(causes=frozenset({dr})))
    graph.add_edge("Sheet1!A1", "Sheet1!B1", provenance=EdgeProvenance(causes=frozenset({dr})))
    graph.set_node_metadata("Sheet1!B1", {"label": "doubled input"})

    class InlineCollapse:
        """Inline B1's body into A1 and delete B1 (no public forwarding)."""

        def project(self, source: DependencyGraph) -> ProjectionResult:
            projected = source.copy()
            node_b = projected.get_node("Sheet1!B1")
            assert node_b is not None
            snapshot = ProjectedNodeSnapshot(
                address="Sheet1!B1",
                sheet=node_b.sheet,
                column=node_b.column,
                row=node_b.row,
                formula=node_b.formula,
                normalized_formula=node_b.normalized_formula,
                value=node_b.value,
                is_target=node_b.is_target,
                is_leaf=node_b.is_leaf,
                metadata=dict(node_b.metadata),
            )
            projected.set_node_formula("Sheet1!A1", "=Sheet1!D1*2+1", "=Sheet1!D1*2+1")
            projected.add_edge(
                "Sheet1!A1",
                "Sheet1!D1",
                provenance=EdgeProvenance(causes=frozenset({dr})),
            )
            projected.remove_node("Sheet1!B1")
            projected.set_node_metadata("Sheet1!A1", {"collapsed_from": ["Sheet1!B1"]})
            manifest = BaseProjectionManifest(
                kind="inline_collapse",
                forwarding_map={},
                retained_to_collapsed_sources={"Sheet1!A1": ("Sheet1!B1",)},
                removed_node_snapshots={"Sheet1!B1": snapshot},
                formula_rewrites=(),
                collapsed_groups=(),
            )
            return ProjectionResult(source, projected, manifest)

    projection = InlineCollapse().project(graph)

    assert "Sheet1!B1" in graph
    assert "Sheet1!B1" not in projection
    assert projection.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!D1"})
    assert projection.manifest.map_to_projected("Sheet1!B1") == "Sheet1!B1"

    assert isinstance(projection.manifest, BaseProjectionManifest)
    snapshot = projection.manifest.removed_node_snapshots["Sheet1!B1"]
    assert snapshot.formula == "=Sheet1!D1*2"
    assert snapshot.metadata["label"] == "doubled input"

    condensed = projection.get_node("Sheet1!A1")
    assert condensed is not None
    assert condensed.metadata["collapsed_from"] == ["Sheet1!B1"]


def test_projection_copy_isolates_projection_mutations() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    for n in (c, b):
        graph.add_node(n)
    graph.add_edge(
        "Sheet1!B1",
        "Sheet1!C1",
        provenance=EdgeProvenance(causes=frozenset({DependencyCause.direct_ref})),
    )
    graph.set_node_metadata("Sheet1!B1", {"label": "source"})

    projected = graph._copy_for_projection()
    projected.set_node_formula("Sheet1!B1", "=1", "=1")
    projected.set_node_metadata("Sheet1!B1", {"label": "projected"})
    projected.remove_node("Sheet1!C1")

    original = graph.get_node("Sheet1!B1")
    assert original is not None
    assert original.normalized_formula == "=Sheet1!C1"
    assert original.metadata["label"] == "source"
    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!C1"})
    assert "Sheet1!C1" in graph


def test_projection_copy_preserves_graph_metadata_fields() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    graph.leaf_classification = {"Sheet1!A1": "input"}
    graph.sheet_order = ["Sheet1", "Sheet2"]
    graph.sheet_bounds = {"Sheet1": (1, 10)}
    graph.named_ranges = {"Rate": ("Sheet1", "A1")}
    graph.named_range_ranges = {"Table": ("Sheet1", "A1", "B2")}
    graph.preparsed_formulas = cast(Any, {"Sheet1!A1": object()})

    graph_structure_fields = {
        "_nodes",
        "_edges",
        "_reverse_edges",
        "_guards",
        "_edge_extra",
        "_hooks",
    }
    metadata_field_names = tuple(
        field.name for field in fields(DependencyGraph) if field.name not in graph_structure_fields
    )
    assert metadata_field_names == (
        "leaf_classification",
        "sheet_order",
        "sheet_bounds",
        "named_ranges",
        "named_range_ranges",
        "preparsed_formulas",
    )

    projected = graph._copy_for_projection()
    for field_name in metadata_field_names:
        original_value = getattr(graph, field_name)
        projected_value = getattr(projected, field_name)
        assert projected_value == original_value
        assert projected_value is not original_value


def test_optimal_projection_does_not_mutate_shared_preparsed_ast() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=SUM(Sheet1!C1,1)", "=SUM(Sheet1!C1,1)")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge(
        "Sheet1!B1",
        "Sheet1!C1",
        provenance=EdgeProvenance(causes=frozenset({dr})),
    )
    ref = "Sheet1!B1"
    formula = "=Sheet1!B1+1"
    span = ((formula.index(ref), formula.index(ref) + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=span,
            direct_sites_normalized=span,
        ),
    )
    ast = parse("=SUM(Sheet1!C1,1)")
    assert isinstance(ast, FunctionCallNode)
    original_args = tuple(ast.args)
    graph.preparsed_formulas = {"=SUM(Sheet1!C1,1)": ast}

    projection = OptimalCompression().project(graph)

    assert "Sheet1!B1" not in projection
    assert graph.preparsed_formulas is not None
    projected_cache = projection.projected_graph.preparsed_formulas
    assert projected_cache is not None
    assert projected_cache is not graph.preparsed_formulas
    assert projected_cache["=SUM(Sheet1!C1,1)"] is ast
    assert graph.preparsed_formulas["=SUM(Sheet1!C1,1)"] is ast
    assert ast.args == list(original_args)


def test_projection_snapshot_and_rewrite_types_are_shared_across_layers() -> None:
    from excel_grapher.grapher import compression as grapher_compression

    assert ProjectedNodeSnapshot is grapher_compression.ProjectedNodeSnapshot
    from excel_grapher.exporter.projection import FormulaRewrite

    assert FormulaRewrite is grapher_compression.FormulaRewrite


def test_dependency_graph_and_projection_satisfy_graph_read_view() -> None:
    from excel_grapher.grapher.graph import DependencyGraph, GraphReadView

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    for n in (c, b):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))

    projection = IdentityTransitCompression().project(graph)

    assert isinstance(graph, GraphReadView)
    assert isinstance(projection, GraphReadView)


def test_set_node_formula_updates_normalized_formula() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    graph.add_node(_make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1"))

    graph.set_node_formula("Sheet1!A1", "=Sheet1!C1", "=Sheet1!C1")
    updated = graph.get_node("Sheet1!A1")
    assert updated is not None
    assert updated.formula == "=Sheet1!C1"
    assert updated.normalized_formula == "=Sheet1!C1"

    graph.set_node_formula("Sheet1!A1", None, None)
    cleared = graph.get_node("Sheet1!A1")
    assert cleared is not None
    assert cleared.formula is None
    assert cleared.normalized_formula is None
