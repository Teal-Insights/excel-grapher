"""Analysis facade for bound catalogs, statements, and classified instance edges.

Visualization and inverted-tree emission share this surface. It does not
import package emission (`emit.py`, `ast_emit.py`, or rung/fusion planning).
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    Statement,
    build_catalog,
)
from excel_grapher.exporter.inverted_tree.deps import (
    CatalogEdges,
    DependenceEdge,
    bind_blank_rects,
    collect_catalog_edges,
    reset_blank_rects,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.blank_ranges import normalize_blank_range_specs
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

__all__ = [
    "BoundSeries",
    "CatalogEdges",
    "DependenceEdge",
    "SemanticCatalogError",
    "SemanticCatalogView",
    "SeriesCatalog",
    "Statement",
    "build_catalog",
    "collect_catalog_edges",
    "load_semantic_catalog",
]


class SemanticCatalogError(ValueError):
    """Catalog or instance-edge analysis could not be completed."""


@dataclass(frozen=True, slots=True)
class SemanticCatalogView:
    """Bound catalog plus classified instance edges for one workbook.

    `catalog` is the statement-partitioned series table. `edges` are the
    instance-level `DependenceEdge`s walked from formula series.
    """

    catalog: SeriesCatalog
    edges: CatalogEdges


def load_semantic_catalog(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    blank_ranges: Iterable[str] | None = None,
) -> SemanticCatalogView:
    """Build the shared catalog and classified edges from graph plus bindings.

    Args:
        graph: Cell-level dependency graph.
        bindings: Validated series bindings.
        workbook: Workbook path used to expand series geometry.
        blank_ranges: Sheet-qualified rectangles omitted from the graph.
            Formula range walks treat them as positional blanks instead of
            fail-closing as unbound series.

    Raises:
        SemanticCatalogError: Catalog expansion or edge classification fail-closed.
    """
    rects = normalize_blank_range_specs(blank_ranges)
    token = bind_blank_rects(rects)
    try:
        try:
            catalog = build_catalog(bindings, workbook=workbook, graph=graph)
            edges = collect_catalog_edges(catalog, graph, blank_rects=rects)
        except InvertedTreeExportError as exc:
            raise SemanticCatalogError(str(exc)) from exc
    finally:
        reset_blank_rects(token)
    return SemanticCatalogView(catalog=catalog, edges=edges)
