"""Shared workbook compression workflow for CLI and examples."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from excel_grapher import create_dependency_graph
from excel_grapher.exporter.projection import (
    BaseProjectionManifest,
    IdentityTransitCompression,
    OptimalCompression,
    ProjectionResult,
    build_similarity_projection_manifest,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.similarity_compression import (
    EmbeddingProvider,
    MockEmbeddingProvider,
    PackingScore,
)

CompressionMethod = Literal["similarity", "optimal", "identity"]


@dataclass(frozen=True)
class CompressionReport:
    """Summary of a graph compression run."""

    method: CompressionMethod
    workbook: Path
    targets: tuple[str, ...]
    original_node_count: int
    projected_node_count: int
    removed_nodes: tuple[str, ...]
    retained_roots: tuple[str, ...]
    manifest_kind: str
    collapsed_groups: tuple[dict[str, Any], ...]
    score: PackingScore | None = None

    @property
    def removed_count(self) -> int:
        """Number of nodes removed by compression."""
        return len(self.removed_nodes)

    def to_dict(self) -> dict[str, Any]:
        """Serialize the report for JSON output."""
        payload: dict[str, Any] = {
            "method": self.method,
            "workbook": str(self.workbook),
            "targets": list(self.targets),
            "original_node_count": self.original_node_count,
            "projected_node_count": self.projected_node_count,
            "removed_count": self.removed_count,
            "removed_nodes": list(self.removed_nodes),
            "retained_roots": list(self.retained_roots),
            "manifest_kind": self.manifest_kind,
            "collapsed_groups": list(self.collapsed_groups),
        }
        if self.score is not None:
            payload["score"] = {
                "final_score": self.score.final_score,
                "total_reduction": self.score.total_reduction,
                "mean_cluster_distance": self.score.mean_cluster_distance,
                "singleton_cluster_fraction": self.score.singleton_cluster_fraction,
            }
        return payload


def build_graph(
    workbook: Path,
    targets: list[str],
    *,
    load_values: bool = True,
) -> DependencyGraph:
    """Build a dependency graph with compression provenance enabled."""
    return create_dependency_graph(
        workbook,
        targets,
        load_values=load_values,
        capture_dependency_provenance=True,
    )


def compress_graph(
    graph: DependencyGraph,
    *,
    method: CompressionMethod = "similarity",
    provider: EmbeddingProvider | None = None,
    preserve: set[str] | None = None,
) -> tuple[ProjectionResult, CompressionReport]:
    """Compress ``graph`` with the requested projection method."""
    if method == "similarity":
        from excel_grapher.grapher.similarity_compression import select_similarity_projection

        embedder = provider or MockEmbeddingProvider()
        selection = select_similarity_projection(
            graph,
            preserve=preserve,
            provider=embedder,
        )
        manifest = build_similarity_projection_manifest(graph, selection.simulation.record)
        projection = ProjectionResult(
            original_graph=graph,
            projected_graph=selection.simulation.projected_graph,
            manifest=manifest,
        )
        score = selection.score
        retained_roots = tuple(group.root for group in selection.packing.groups)
    elif method == "optimal":
        projection = OptimalCompression(preserve=preserve).project(graph)
        score = None
        retained_roots = tuple(
            key for key in graph.target_keys() if key in projection.projected_graph
        )
    elif method == "identity":
        projection = IdentityTransitCompression().project(graph)
        score = None
        retained_roots = tuple(
            key for key in graph.target_keys() if key in projection.projected_graph
        )
    else:
        raise ValueError(f"Unsupported compression method: {method!r}")

    manifest = projection.manifest
    if not isinstance(manifest, BaseProjectionManifest):
        raise TypeError("Expected BaseProjectionManifest from compression projection")

    removed_nodes = tuple(sorted(key for key in graph if key not in projection.projected_graph))
    report = CompressionReport(
        method=method,
        workbook=Path("<in-memory>"),
        targets=tuple(graph.target_keys()),
        original_node_count=len(graph),
        projected_node_count=len(projection.projected_graph),
        removed_nodes=removed_nodes,
        retained_roots=retained_roots,
        manifest_kind=manifest.kind,
        collapsed_groups=tuple(group.to_dict() for group in manifest.collapsed_groups),
        score=score,
    )
    return projection, report


def compress_workbook(
    workbook: Path,
    targets: list[str],
    *,
    method: CompressionMethod = "similarity",
    provider: EmbeddingProvider | None = None,
    preserve: set[str] | None = None,
    load_values: bool = True,
) -> tuple[ProjectionResult, CompressionReport]:
    """Build a graph from ``workbook`` and compress it."""
    graph = build_graph(workbook, targets, load_values=load_values)
    projection, report = compress_graph(
        graph,
        method=method,
        provider=provider,
        preserve=preserve,
    )
    report = CompressionReport(
        method=report.method,
        workbook=workbook,
        targets=tuple(targets),
        original_node_count=report.original_node_count,
        projected_node_count=report.projected_node_count,
        removed_nodes=report.removed_nodes,
        retained_roots=report.retained_roots,
        manifest_kind=report.manifest_kind,
        collapsed_groups=report.collapsed_groups,
        score=report.score,
    )
    return projection, report


def format_report_text(report: CompressionReport) -> str:
    """Render a human-readable compression report."""
    lines = [
        f"Method: {report.method}",
        f"Workbook: {report.workbook}",
        (
            f"Nodes: {report.original_node_count} -> {report.projected_node_count} "
            f"({report.removed_count} removed)"
        ),
    ]
    if report.score is not None:
        lines.append(
            "Score: "
            f"{report.score.final_score:.3f} "
            f"(reduction={report.score.total_reduction}, "
            f"mean_cluster_distance={report.score.mean_cluster_distance:.4f}, "
            f"singleton_fraction={report.score.singleton_cluster_fraction:.2f})"
        )
    lines.append("")
    lines.append("Retained roots:")
    for root in report.retained_roots:
        lines.append(f"  {root}")
    lines.append("")
    lines.append(f"Collapsed groups ({len(report.collapsed_groups)}):")
    for group in report.collapsed_groups:
        retained = group["retained"]
        sources = ", ".join(group["collapsed_sources"])
        lines.append(f"  {retained} <- [{sources}]")
    return "\n".join(lines)
