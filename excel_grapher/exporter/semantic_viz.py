"""Compressed statement-graph visualization payload and HTML writer."""

from __future__ import annotations

import json
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from importlib import resources
from pathlib import Path
from typing import Any, Literal

from excel_grapher.exporter.semantic_catalog import (
    SemanticCatalogView,
    load_semantic_catalog,
)
from excel_grapher.exporter.semantic_graph import (
    StatementGraph,
    build_statement_graph,
    jsonable_scalar,
    schedule_edges,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.viz_layout import (
    DEFAULT_RANK_PULL,
    RANK_PULL_MODES,
    RankPull,
    VizTier,
    clustered_force_layout,
    input_depths,
    viz_layout_js_bootstrap,
    viz_tier,
)
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

SEMANTIC_VIZ_PAYLOAD_VERSION = 2
SEMANTIC_VIZ_OVERLAY_ID = "webviz.statement_graph"
SEMANTIC_VIZ_CELL_SAMPLE = 8
SEMANTIC_VIZ_CLUSTER_BY = "series"

__all__ = [
    "SEMANTIC_VIZ_CELL_SAMPLE",
    "SEMANTIC_VIZ_CLUSTER_BY",
    "SEMANTIC_VIZ_OVERLAY_ID",
    "SEMANTIC_VIZ_PAYLOAD_VERSION",
    "SemanticVizPayload",
    "semantic_viz_primitive_count",
    "semantic_viz_tier",
    "serialize_semantic_viz_json",
    "statement_graph_layout",
    "to_semantic_viz_payload",
    "write_semantic_viz_html",
]


def semantic_viz_primitive_count(*, statement_count: int, bundle_count: int) -> int:
    """Return statement + twice-bundle size used to pick the viewer tier."""
    return statement_count + 2 * bundle_count


def semantic_viz_tier(*, statement_count: int, bundle_count: int) -> VizTier:
    """Return the shared size tier for a statement graph.

    `small` and `medium` draw labeled boxes; `large` opens on the cluster
    overview and draws dots when zoomed in.
    """
    return viz_tier(
        semantic_viz_primitive_count(statement_count=statement_count, bundle_count=bundle_count)
    )


def statement_graph_layout(
    graph: StatementGraph, *, rank_pull: RankPull = DEFAULT_RANK_PULL
) -> tuple[tuple[tuple[float, float], ...], tuple[int, ...]]:
    """Lay out a statement graph with the shared clustered force layout.

    Clusters are series. Depth is input depth over the schedule edges
    (`schedule_edges`), so the pull puts inputs at the top.

    Returns:
        `(positions, depths)`, one entry per `graph.nodes`.
    """
    nodes = graph.nodes
    n = len(nodes)
    index = {node.statement_id: i for i, node in enumerate(nodes)}
    edges = sorted(
        {
            (index[b.consumer_id], index[b.producer_id])
            for b in graph.bundles
            if b.consumer_id in index and b.producer_id in index and b.consumer_id != b.producer_id
        }
    )
    depths = tuple(input_depths(n, schedule_edges(nodes, graph.bundles)))
    pos = clustered_force_layout(
        n,
        edges,
        [node.series_id for node in nodes],
        depths=depths,
        rank_pull=rank_pull,
    )
    return tuple((float(x), float(y)) for x, y in pos.tolist()), depths


@dataclass(frozen=True, slots=True)
class SemanticVizPayload:
    """JSON-serializable statement-graph visualization.

    Distinct from cell-level `LightweightVizPayload`. Nodes are statements,
    not cells. `ranks` on the graph come from the distance-zero residual;
    positions come from the shared clustered force layout (clusters are
    series) with the vertical `rank_pull` towards input depth.
    """

    version: int
    graph: StatementGraph
    positions: tuple[tuple[float, float], ...]
    depths: tuple[int, ...]
    rank_pull: RankPull = DEFAULT_RANK_PULL
    annotations: Mapping[str, Any] | None = None

    @property
    def tier(self) -> VizTier:
        """Shared size tier of this graph."""
        return semantic_viz_tier(
            statement_count=len(self.graph.nodes), bundle_count=len(self.graph.bundles)
        )

    def to_dict(self, *, cell_sample: int | None = None) -> dict[str, Any]:
        """Return a JSON-serializable mapping.

        Args:
            cell_sample: If set, keep only the first `cell_sample` addresses on
                each node and the first `cell_sample` instance edges on each
                bundle. The HTML viewer uses this cap; omit it for a full dump
                (`--json`). `cell_count` and `instance_edge_count` stay the
                true sizes.
        """
        g = self.graph
        return {
            "version": self.version,
            "kind": "statement_graph",
            "overlay_id": SEMANTIC_VIZ_OVERLAY_ID,
            "layout": {
                "tier": self.tier,
                "rank_pull": self.rank_pull,
                "cluster_by": SEMANTIC_VIZ_CLUSTER_BY,
            },
            "stats": {
                "cell_count": g.stats.cell_count,
                "statement_count": g.stats.statement_count,
                "instance_edge_count": g.stats.instance_edge_count,
                "bundle_count": g.stats.bundle_count,
                "unbound_cell_count": g.stats.unbound_cell_count,
                "heterogeneous_partition_pair_count": (g.stats.heterogeneous_partition_pair_count),
            },
            "nodes": [
                {
                    "id": node.statement_id,
                    "series_id": node.series_id,
                    "shape_key": node.shape_key,
                    "start": node.start,
                    "stop": node.stop,
                    "cell_count": node.cell_count,
                    "cells": (
                        list(node.cells) if cell_sample is None else list(node.cells[:cell_sample])
                    ),
                    "cell_partitions": [
                        [jsonable_scalar(value) for value in part]
                        for part in (
                            node.cell_partitions
                            if cell_sample is None
                            else node.cell_partitions[:cell_sample]
                        )
                    ],
                    "partitions": [
                        [jsonable_scalar(value) for value in part] for part in node.partitions
                    ],
                    "direction": node.direction,
                    "sheet": node.sheet,
                    "is_remainder": node.is_remainder,
                    "labels": node.labels.to_dict(),
                    "rank": g.ranks[i],
                    "depth": self.depths[i],
                    "x": self.positions[i][0],
                    "y": self.positions[i][1],
                }
                for i, node in enumerate(g.nodes)
            ],
            "bundles": [
                {
                    "consumer_id": bundle.consumer_id,
                    "producer_id": bundle.producer_id,
                    "access": bundle.access,
                    "distance": bundle.distance,
                    "coeff": bundle.coeff,
                    "offset": bundle.offset,
                    "guarded": bundle.guarded,
                    "instance_edge_count": bundle.instance_edge_count,
                    "partitions": [
                        [jsonable_scalar(v) for v in part] for part in bundle.partitions
                    ],
                    "representative_consumer_cell": bundle.representative_consumer_cell,
                    "representative_producer_cell": bundle.representative_producer_cell,
                    "instance_edges": [
                        {
                            "consumer_cell": consumer,
                            "producer_cell": producer,
                        }
                        for consumer, producer in (
                            bundle.instance_edges
                            if cell_sample is None
                            else bundle.instance_edges[:cell_sample]
                        )
                    ],
                    "discharged": bundle.distance > 0 or bundle.access == "shift",
                }
                for bundle in g.bundles
            ],
            "remainder_sample": list(g.remainder_sample),
            "heterogeneous_pairs": [list(pair) for pair in g.heterogeneous_pairs],
            "annotations": dict(self.annotations) if self.annotations else {},
        }


def to_semantic_viz_payload(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    view: SemanticCatalogView | None = None,
    blank_ranges: Iterable[str] | None = None,
    rank_pull: RankPull = DEFAULT_RANK_PULL,
) -> SemanticVizPayload:
    """Build a statement-graph payload from a cell graph plus series bindings.

    Does not construct cell-level `LightweightVizCore` and does not import
    inverted-tree emission. Catalog analysis fail-closes as `SemanticCatalogError`.

    Args:
        graph: Cell-level dependency graph.
        bindings: Validated series bindings.
        workbook: Workbook path used to expand series geometry.
        view: Precomputed catalog. When omitted, `load_semantic_catalog` runs.
        blank_ranges: Sheet-qualified rectangles omitted from the graph. Forwarded
            to catalog edge classification when `view` is omitted.
        rank_pull: Vertical pull towards input depth: `"none"`, `"between"`
            series clusters (default), or `"everywhere"`.

    Raises:
        SemanticCatalogError: Catalog or instance-edge analysis failed.
        ValueError: Unknown `rank_pull`.
    """
    if rank_pull not in RANK_PULL_MODES:
        raise ValueError(f"rank_pull must be one of {RANK_PULL_MODES}, got {rank_pull!r}")
    snapshot = (
        view
        if view is not None
        else load_semantic_catalog(graph, bindings, workbook=workbook, blank_ranges=blank_ranges)
    )
    statement_graph = build_statement_graph(snapshot, graph)
    positions, depths = statement_graph_layout(statement_graph, rank_pull=rank_pull)
    return SemanticVizPayload(
        version=SEMANTIC_VIZ_PAYLOAD_VERSION,
        graph=statement_graph,
        positions=positions,
        depths=depths,
        rank_pull=rank_pull,
    )


def serialize_semantic_viz_json(
    payload: SemanticVizPayload,
    *,
    cell_sample: int | None = SEMANTIC_VIZ_CELL_SAMPLE,
) -> str:
    """Serialize `payload` as compact JSON.

    Args:
        payload: Statement-graph visualization to encode.
        cell_sample: Address cap per node and per bundle for the HTML viewer.
            Pass `None` to keep every bound cell and instance edge (same as
            `SemanticVizPayload.to_dict()`).
    """
    return json.dumps(
        payload.to_dict(cell_sample=cell_sample),
        separators=(",", ":"),
        ensure_ascii=False,
    )


def write_semantic_viz_html(
    payload: SemanticVizPayload,
    path: Path | str,
    *,
    title: str = "Statement dependency graph",
    data_mode: Literal["inline", "sidecar", "auto"] = "inline",
    template_path: Path | str | None = None,
) -> None:
    """Write a standalone HTML viewer for a statement-graph payload.

    The viewer opens on the precomputed clustered force layout. Behavior
    follows the shared size tier: `small` graphs relayout in the browser when
    Cluster by or Vertical pull changes; `medium` graphs relayout one selected
    cluster or neighborhood; `large` graphs open on the cluster overview,
    expand into dots when zoomed in, and relayout one selected cluster.
    """
    if payload.version != SEMANTIC_VIZ_PAYLOAD_VERSION:
        raise ValueError(f"Unsupported semantic viz payload version: {payload.version}")
    out = Path(path)
    json_payload = serialize_semantic_viz_json(payload)
    sidecar_name: str | None = None
    if data_mode == "sidecar":
        sidecar_name = out.with_suffix(".viz.json").name
        out.parent.mkdir(parents=True, exist_ok=True)
        (out.parent / sidecar_name).write_text(json_payload, encoding="utf-8")
        json_payload = None
    pkg = resources.files(__package__ or __name__)
    if template_path is None:
        tpl = pkg.joinpath("semantic_viz_template.html").read_text(encoding="utf-8")
    else:
        tpl = Path(template_path).read_text(encoding="utf-8")
    layout_js = pkg.joinpath("semantic_viz_layout.js").read_text(encoding="utf-8")
    bootstrap = (
        f"window.__VIZ_DATA__ = {json_payload};"
        if json_payload is not None
        else "window.__VIZ_DATA__ = null;"
    )
    sidecar_js = (
        f"window.__VIZ_DATA_URL__ = {json.dumps(sidecar_name)};"
        if json_payload is None
        else "window.__VIZ_DATA_URL__ = null;"
    )
    html = (
        tpl.replace("__TITLE__", title)
        .replace("/*__BOOTSTRAP__*/", bootstrap)
        .replace("/*__SIDECAR__*/", sidecar_js)
        .replace("/*__VIZ_FORCE_JS__*/", viz_layout_js_bootstrap())
        .replace("/*__LAYOUT_JS__*/", layout_js)
    )
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
