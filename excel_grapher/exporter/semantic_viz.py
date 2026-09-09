"""Compressed statement-graph visualization payload and HTML writer."""

from __future__ import annotations

import json
from collections.abc import Mapping
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
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

SEMANTIC_VIZ_PAYLOAD_VERSION = 1
SEMANTIC_VIZ_OVERLAY_ID = "webviz.statement_graph"
SEMANTIC_VIZ_CELL_SAMPLE = 8

__all__ = [
    "SEMANTIC_VIZ_CELL_SAMPLE",
    "SEMANTIC_VIZ_OVERLAY_ID",
    "SEMANTIC_VIZ_PAYLOAD_VERSION",
    "SemanticVizPayload",
    "serialize_semantic_viz_json",
    "to_semantic_viz_payload",
    "write_semantic_viz_html",
]


@dataclass(frozen=True, slots=True)
class SemanticVizPayload:
    """JSON-serializable statement-graph visualization.

    Distinct from cell-level `LightweightVizPayload`. Nodes are statements,
    not cells. Layout ranks come from the distance-zero residual, not Louvain.
    """

    version: int
    graph: StatementGraph
    annotations: Mapping[str, Any] | None = None

    def to_dict(self, *, cell_sample: int | None = None) -> dict[str, Any]:
        """Return a JSON-serializable mapping.

        Args:
            cell_sample: If set, keep only the first `cell_sample` addresses on
                each node. The HTML viewer uses this cap; omit it for a full
                dump (`--json`). `cell_count` is always the true statement size.
        """
        g = self.graph
        return {
            "version": self.version,
            "kind": "statement_graph",
            "overlay_id": SEMANTIC_VIZ_OVERLAY_ID,
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
                    "direction": node.direction,
                    "is_remainder": node.is_remainder,
                    "rank": g.ranks[i],
                    "x": g.positions[i][0],
                    "y": g.positions[i][1],
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
) -> SemanticVizPayload:
    """Build a statement-graph payload from a cell graph plus series bindings.

    Does not construct cell-level `LightweightVizCore` and does not import
    inverted-tree emission. Catalog analysis fail-closes as `SemanticCatalogError`.

    Raises:
        SemanticCatalogError: Catalog or instance-edge analysis failed.
    """
    snapshot = (
        view if view is not None else load_semantic_catalog(graph, bindings, workbook=workbook)
    )
    statement_graph = build_statement_graph(snapshot, graph)
    return SemanticVizPayload(version=SEMANTIC_VIZ_PAYLOAD_VERSION, graph=statement_graph)


def serialize_semantic_viz_json(
    payload: SemanticVizPayload,
    *,
    cell_sample: int | None = SEMANTIC_VIZ_CELL_SAMPLE,
) -> str:
    """Serialize `payload` as compact JSON.

    Args:
        payload: Statement-graph visualization to encode.
        cell_sample: Address cap per node for the HTML viewer. Pass `None` to
            keep every bound cell (same as `SemanticVizPayload.to_dict()`).
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
    """Write a standalone HTML viewer for a statement-graph payload."""
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
    if template_path is None:
        tpl = (
            resources.files(__package__ or __name__)
            .joinpath("semantic_viz_template.html")
            .read_text(encoding="utf-8")
        )
    else:
        tpl = Path(template_path).read_text(encoding="utf-8")
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
    )
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
