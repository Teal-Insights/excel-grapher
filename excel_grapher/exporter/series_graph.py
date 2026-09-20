"""Bindings-backed constructor for the series-level dependency quotient."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from excel_grapher.exporter.inverted_tree.catalog import SeriesCatalog, build_catalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.series_graph import SeriesGraph
from excel_grapher.grapher.series_graph import to_series_graph as collapse_series_graph
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

from .semantic_catalog import SemanticCatalogError

__all__ = ["to_series_graph"]


def to_series_graph(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    catalog: SeriesCatalog | None = None,
    include_self_loops: bool = False,
    blank_ranges: Iterable[str] | None = None,
) -> SeriesGraph:
    """Collapse a cell graph to one node per bound series.

    This is the series-level analogue of `to_sheet_graph`: cell edges between
    the same pair of series become one directed edge labeled with the
    underlying cell-edge count. Intra-series edges are dropped unless
    `include_self_loops=True`. The quotient is not a schedule DAG and may
    contain cycles even when the cell graph does not.

    Args:
        graph: Cell-level dependency graph.
        bindings: Validated series bindings.
        workbook: Workbook path used to expand series geometry when `catalog`
            is omitted.
        catalog: Precomputed occupancy catalog. When omitted, `build_catalog`
            runs.
        include_self_loops: If `True`, keep intra-series cell edges as a
            self-loop per series.
        blank_ranges: Sheet-qualified rectangles omitted from the graph.
            Forwarded to catalog expansion when `catalog` is omitted.

    Returns:
        Immutable series graph. Node order follows bindings order.

    Raises:
        SemanticCatalogError: Catalog expansion fail-closed.
    """
    try:
        occupancy = (
            catalog
            if catalog is not None
            else build_catalog(
                bindings,
                workbook=workbook,
                graph=graph,
                blank_ranges=list(blank_ranges) if blank_ranges is not None else None,
            )
        )
    except InvertedTreeExportError as exc:
        raise SemanticCatalogError(str(exc)) from exc
    return collapse_series_graph(graph, occupancy, include_self_loops=include_self_loops)
