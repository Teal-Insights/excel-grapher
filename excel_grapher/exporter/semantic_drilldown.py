"""Cell-level drilldown from a statement-graph series or statement."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from excel_grapher.core.address_keys import CanonicalAddress
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, Statement
from excel_grapher.exporter.inverted_tree.deps import DependenceEdge
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.semantic_catalog import SemanticCatalogError, SemanticCatalogView
from excel_grapher.exporter.semantic_graph import jsonable_scalar
from excel_grapher.grapher.formula_label import display_formula
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.types import Scalar

__all__ = [
    "DrilldownCell",
    "SeriesDrilldown",
    "drilldown_series",
    "drilldown_statement",
]


@dataclass(frozen=True, slots=True)
class DrilldownCell:
    """One cell in a series drilldown, with formula text and catalog identity."""

    address: CanonicalAddress
    series_id: str | None
    statement_id: str | None
    shape_key: str | None
    index: int | None
    formula: str | None
    value: object
    key: tuple[tuple[str, Scalar], ...]
    in_series: bool

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this cell."""
        return {
            "address": self.address,
            "series_id": self.series_id,
            "statement_id": self.statement_id,
            "shape_key": self.shape_key,
            "index": self.index,
            "formula": self.formula,
            "value": _jsonable_value(self.value),
            "key": {name: jsonable_scalar(value) for name, value in self.key},
            "in_series": self.in_series,
        }


@dataclass(frozen=True, slots=True)
class SeriesDrilldown:
    """Cells and classified adjacency for one series, or one statement inside it.

    `cells` are the selected series members in catalog order. `neighbors` are
    one-hop producer/consumer cells outside that selection. `in_series` is True
    when the cell belongs to the drilled series, including same-series neighbors
    of a single-statement drilldown. `edges` are instance edges that touch the
    selection. Formula text lives here, not on statement nodes.
    """

    series_id: str
    statement_id: str | None
    statements: tuple[Statement, ...]
    cells: tuple[DrilldownCell, ...]
    neighbors: tuple[DrilldownCell, ...]
    edges: tuple[DependenceEdge, ...]

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this drilldown."""
        return {
            "series_id": self.series_id,
            "statement_id": self.statement_id,
            "statements": [
                {
                    "id": stmt.statement_id,
                    "shape_key": stmt.shape_key,
                    "start": stmt.start,
                    "stop": stmt.stop,
                    "cell_count": len(stmt.cells),
                }
                for stmt in self.statements
            ],
            "cells": [cell.to_dict() for cell in self.cells],
            "neighbors": [cell.to_dict() for cell in self.neighbors],
            "edges": [_edge_dict(edge) for edge in self.edges],
        }

    def to_networkx(self):
        """Return a NetworkX MultiDiGraph of drilldown cells and instance edges."""
        return _drilldown_to_networkx(self)


def drilldown_series(
    view: SemanticCatalogView,
    graph: DependencyGraph,
    series_id: str,
) -> SeriesDrilldown:
    """Return cells and adjacency for every statement in `series_id`.

    Args:
        view: Statement-partitioned catalog and classified instance edges.
        graph: Cell-level graph used for formula text and values.
        series_id: Bound series to expand.

    Raises:
        SemanticCatalogError: `series_id` is not in the catalog.
    """
    series = _require_series(view.catalog, series_id)
    return _drilldown(view, graph, series, statement=None)


def drilldown_statement(
    view: SemanticCatalogView,
    graph: DependencyGraph,
    statement_id: str,
) -> SeriesDrilldown:
    """Return cells and adjacency for one statement.

    Args:
        view: Statement-partitioned catalog and classified instance edges.
        graph: Cell-level graph used for formula text and values.
        statement_id: Statement to expand.

    Raises:
        SemanticCatalogError: `statement_id` is not in the catalog.
    """
    found = _find_statement(view.catalog, statement_id)
    if found is None:
        raise SemanticCatalogError(f"unknown statement {statement_id!r}")
    series, statement = found
    return _drilldown(view, graph, series, statement=statement)


def _require_series(catalog: SeriesCatalog, series_id: str) -> BoundSeries:
    try:
        return catalog.get(series_id)
    except InvertedTreeExportError as exc:
        raise SemanticCatalogError(str(exc)) from exc


def _find_statement(
    catalog: SeriesCatalog, statement_id: str
) -> tuple[BoundSeries, Statement] | None:
    for series_id in catalog.order:
        series = catalog.get(series_id)
        for statement in series.statements:
            if statement.statement_id == statement_id:
                return series, statement
    return None


def _statement_by_cell(catalog: SeriesCatalog) -> dict[CanonicalAddress, Statement]:
    mapping: dict[CanonicalAddress, Statement] = {}
    for series_id in catalog.order:
        series = catalog.get(series_id)
        for statement in series.statements:
            for cell in statement.cells:
                mapping[cell] = statement
    return mapping


def _drilldown(
    view: SemanticCatalogView,
    graph: DependencyGraph,
    series: BoundSeries,
    *,
    statement: Statement | None,
) -> SeriesDrilldown:
    selected_cells = statement.cells if statement is not None else series.cells
    selected = set(selected_cells)
    statements = (statement,) if statement is not None else series.statements
    if statement is None:
        edges = tuple(
            edge
            for edge in view.edges.edges
            if edge.consumer_id == series.series_id or edge.producer_id == series.series_id
        )
    else:
        edges = tuple(
            edge
            for edge in view.edges.edges
            if edge.consumer_cell in selected or edge.producer_cell in selected
        )

    stmt_by_cell = _statement_by_cell(view.catalog)
    cells = tuple(
        _cell_record(
            address,
            view.catalog,
            graph,
            stmt_by_cell,
            drilled_series_id=series.series_id,
        )
        for address in selected_cells
    )
    neighbor_addresses = sorted(
        {
            endpoint
            for edge in edges
            for endpoint in (edge.consumer_cell, edge.producer_cell)
            if endpoint not in selected
        }
    )
    neighbors = tuple(
        _cell_record(
            address,
            view.catalog,
            graph,
            stmt_by_cell,
            drilled_series_id=series.series_id,
        )
        for address in neighbor_addresses
    )
    return SeriesDrilldown(
        series_id=series.series_id,
        statement_id=None if statement is None else statement.statement_id,
        statements=statements,
        cells=cells,
        neighbors=neighbors,
        edges=edges,
    )


def _cell_record(
    address: CanonicalAddress,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    stmt_by_cell: Mapping[CanonicalAddress, Statement],
    *,
    drilled_series_id: str,
) -> DrilldownCell:
    owner = catalog.series_for(address)
    covering = stmt_by_cell.get(address)
    node = graph.get_node(address)
    point = owner.key_point_for(address) if owner is not None else None
    return DrilldownCell(
        address=address,
        series_id=None if owner is None else owner.series_id,
        statement_id=None if covering is None else covering.statement_id,
        shape_key=None if covering is None else covering.shape_key,
        index=None if owner is None else owner.index_of(address),
        formula=None if node is None else display_formula(node),
        value=None if node is None else node.value,
        key=() if point is None else point.items,
        in_series=owner is not None and owner.series_id == drilled_series_id,
    )


def _edge_dict(edge: DependenceEdge) -> dict[str, Any]:
    return {
        "consumer_id": edge.consumer_id,
        "producer_id": edge.producer_id,
        "consumer_cell": edge.consumer_cell,
        "producer_cell": edge.producer_cell,
        "access": edge.access,
        "distance": edge.distance,
        "coeff": edge.coeff,
        "offset": edge.offset,
        "guarded": edge.guarded,
    }


def _jsonable_value(value: object) -> object:
    if isinstance(value, datetime):
        return value.isoformat()
    return value


def _edge_key(edge: DependenceEdge) -> str:
    return f"{edge.access}:{edge.distance}:{edge.coeff}:{edge.offset}:{int(edge.guarded)}"


def _drilldown_to_networkx(drilldown: SeriesDrilldown):
    try:
        import networkx as nx
    except Exception as exc:  # pragma: no cover
        raise ImportError(
            "networkx is not installed; add it to use SeriesDrilldown.to_networkx()"
        ) from exc

    nx_graph = nx.MultiDiGraph()
    for cell in (*drilldown.cells, *drilldown.neighbors):
        nx_graph.add_node(cell.address, **cell.to_dict())
    for edge in drilldown.edges:
        nx_graph.add_edge(
            edge.consumer_cell,
            edge.producer_cell,
            key=_edge_key(edge),
            **_edge_dict(edge),
        )
    return nx_graph
