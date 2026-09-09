"""Analysis facade for bound catalogs, statements, and classified instance edges.

Visualization and inverted-tree emission share this surface. It does not
import package emission (`emit.py`, `ast_emit.py`, or rung/fusion planning).
"""

from __future__ import annotations

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
    collect_catalog_edges,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
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
) -> SemanticCatalogView:
    """Build the shared catalog and classified edges from graph plus bindings.

    Raises:
        SemanticCatalogError: Catalog expansion or edge classification fail-closed.
    """
    try:
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
        edges = collect_catalog_edges(catalog, graph)
    except InvertedTreeExportError as exc:
        raise SemanticCatalogError(str(exc)) from exc
    return SemanticCatalogView(catalog=catalog, edges=edges)
