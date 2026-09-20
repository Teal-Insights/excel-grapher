"""Cell-level drilldown from a statement-graph series or statement."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, Statement
from excel_grapher.exporter.inverted_tree.deps import AccessClass, DependenceEdge
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.semantic_catalog import SemanticCatalogError, SemanticCatalogView
from excel_grapher.exporter.semantic_graph import (
    REMAINDER_STATEMENT_ID,
    jsonable_cell_value,
    jsonable_scalar,
)
from excel_grapher.grapher.formula_label import display_formula
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.types import Scalar

__all__ = [
    "DrilldownCell",
    "DrilldownEdge",
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
            "value": jsonable_cell_value(self.value),
            "key": {name: jsonable_scalar(value) for name, value in self.key},
            "in_series": self.in_series,
        }


@dataclass(frozen=True, slots=True)
class DrilldownEdge:
    """One classified instance edge with statement-graph endpoints."""

    consumer_id: str
    producer_id: str
    consumer_statement_id: str | None
    producer_statement_id: str | None
    consumer_cell: CanonicalAddress
    producer_cell: CanonicalAddress
    distance: int
    access: AccessClass
    coeff: int | None = None
    offset: int | None = None
    guarded: bool = False

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this edge."""
        return {
            "consumer_id": self.consumer_id,
            "producer_id": self.producer_id,
            "consumer_statement_id": self.consumer_statement_id,
            "producer_statement_id": self.producer_statement_id,
            "consumer_cell": self.consumer_cell,
            "producer_cell": self.producer_cell,
            "access": self.access,
            "distance": self.distance,
            "coeff": self.coeff,
            "offset": self.offset,
            "guarded": self.guarded,
        }


@dataclass(frozen=True, slots=True)
class SeriesDrilldown:
    """Cells and classified adjacency for one series, or one statement inside it.

    `cells` are the selected series members in catalog order. `neighbors` are
    one-hop producer/consumer cells outside that selection. `in_series` is True
    when the cell belongs to the drilled series, including same-series neighbors
    of a single-statement drilldown. `edges` are instance edges that touch the
    selection, with statement ids for joining to the statement graph. Formula
    text lives here, not on statement nodes. Unbound remainder cells have no
    classified catalog edges.
    """

    series_id: str
    statement_id: str | None
    statements: tuple[Statement, ...]
    cells: tuple[DrilldownCell, ...]
    neighbors: tuple[DrilldownCell, ...]
    edges: tuple[DrilldownEdge, ...]

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
            "edges": [edge.to_dict() for edge in self.edges],
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

    `statement_id` `__unbound__` returns unbound graph cells. Classified
    catalog edges never include those cells, so `edges` is empty.

    Args:
        view: Statement-partitioned catalog and classified instance edges.
        graph: Cell-level graph used for formula text and values.
        statement_id: Statement to expand, or `__unbound__`.

    Raises:
        SemanticCatalogError: `statement_id` is not in the catalog and the
            graph has no unbound remainder cells.
    """
    if statement_id == REMAINDER_STATEMENT_ID:
        remainder = _drilldown_remainder(view, graph)
        if remainder.cells:
            return remainder
        raise SemanticCatalogError(f"unknown statement {statement_id!r}")
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


def _covering_statement(catalog: SeriesCatalog, address: CanonicalAddress) -> Statement | None:
    owner = catalog.series_for(address)
    if owner is None:
        return None
    index = owner.index_of(address)
    if index is None:
        return None
    for statement in owner.statements:
        if statement.start <= index < statement.stop:
            return statement
    return None


def _unique_edges(
    outgoing: tuple[DependenceEdge, ...],
    incoming: tuple[DependenceEdge, ...],
) -> tuple[DependenceEdge, ...]:
    seen: dict[DependenceEdge, None] = {}
    for edge in (*outgoing, *incoming):
        seen.setdefault(edge, None)
    return tuple(seen)


def _series_touching_edges(
    view: SemanticCatalogView,
    series_id: str,
    selected: set[CanonicalAddress] | None,
) -> tuple[DependenceEdge, ...]:
    candidates = _unique_edges(
        view.edges.by_consumer.get(series_id, ()),
        view.edges.by_producer.get(series_id, ()),
    )
    if selected is None:
        return candidates
    return tuple(
        edge
        for edge in candidates
        if edge.consumer_cell in selected or edge.producer_cell in selected
    )


def _as_drilldown_edge(edge: DependenceEdge, catalog: SeriesCatalog) -> DrilldownEdge:
    consumer = _covering_statement(catalog, edge.consumer_cell)
    producer = _covering_statement(catalog, edge.producer_cell)
    return DrilldownEdge(
        consumer_id=edge.consumer_id,
        producer_id=edge.producer_id,
        consumer_statement_id=None if consumer is None else consumer.statement_id,
        producer_statement_id=None if producer is None else producer.statement_id,
        consumer_cell=edge.consumer_cell,
        producer_cell=edge.producer_cell,
        distance=edge.distance,
        access=edge.access,
        coeff=edge.coeff,
        offset=edge.offset,
        guarded=edge.guarded,
    )


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
    raw_edges = _series_touching_edges(
        view,
        series.series_id,
        None if statement is None else selected,
    )
    edges = tuple(_as_drilldown_edge(edge, view.catalog) for edge in raw_edges)
    cells = tuple(
        _cell_record(
            address,
            view.catalog,
            graph,
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


def _unbound_addresses(
    view: SemanticCatalogView, graph: DependencyGraph
) -> tuple[CanonicalAddress, ...]:
    unbound = [
        as_canonical(key)
        for key in graph.keys(order="workbook")
        if key not in view.catalog.address_to_id
    ]
    unbound.sort()
    return tuple(unbound)


def _drilldown_remainder(view: SemanticCatalogView, graph: DependencyGraph) -> SeriesDrilldown:
    cells = tuple(
        _cell_record(
            address,
            view.catalog,
            graph,
            drilled_series_id=REMAINDER_STATEMENT_ID,
            remainder=True,
        )
        for address in _unbound_addresses(view, graph)
    )
    return SeriesDrilldown(
        series_id=REMAINDER_STATEMENT_ID,
        statement_id=REMAINDER_STATEMENT_ID,
        statements=(),
        cells=cells,
        neighbors=(),
        edges=(),
    )


def _cell_record(
    address: CanonicalAddress,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    drilled_series_id: str,
    remainder: bool = False,
) -> DrilldownCell:
    owner = catalog.series_for(address)
    covering = _covering_statement(catalog, address)
    node = graph.get_node(address)
    point = owner.key_point_for(address) if owner is not None else None
    statement_id = (
        REMAINDER_STATEMENT_ID
        if remainder
        else (None if covering is None else covering.statement_id)
    )
    return DrilldownCell(
        address=address,
        series_id=None if owner is None else owner.series_id,
        statement_id=statement_id,
        shape_key=None if covering is None else covering.shape_key,
        index=None if owner is None else owner.index_of(address),
        formula=None if node is None else display_formula(node),
        value=None if node is None else node.value,
        key=() if point is None else point.items,
        in_series=owner is not None and owner.series_id == drilled_series_id,
    )


def _edge_key(edge: DrilldownEdge) -> str:
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
            **edge.to_dict(),
        )
    return nx_graph
