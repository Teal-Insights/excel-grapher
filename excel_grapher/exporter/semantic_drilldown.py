"""Cell-level drilldown from a statement-graph series, statement, or bundle."""

from __future__ import annotations

from collections.abc import Sequence
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
    "BundleDrilldown",
    "DrilldownCell",
    "DrilldownEdge",
    "DrilldownStatement",
    "SeriesDrilldown",
    "drilldown_bundle",
    "drilldown_partition",
    "drilldown_series",
    "drilldown_statement",
]


@dataclass(frozen=True, slots=True)
class DrilldownStatement:
    """Statement identity for a series drilldown, without catalog types."""

    statement_id: str
    shape_key: str | None
    start: int
    stop: int
    cell_count: int

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this statement."""
        return {
            "id": self.statement_id,
            "shape_key": self.shape_key,
            "start": self.start,
            "stop": self.stop,
            "cell_count": self.cell_count,
        }


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
    partition: tuple[Scalar, ...] = ()

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
            "partition": [jsonable_scalar(value) for value in self.partition],
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
    partition: tuple[Scalar, ...] = ()

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
            "partition": [jsonable_scalar(value) for value in self.partition],
        }


@dataclass(frozen=True, slots=True)
class SeriesDrilldown:
    """Cells and classified adjacency for one series, statement, or partition.

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
    statements: tuple[DrilldownStatement, ...]
    cells: tuple[DrilldownCell, ...]
    neighbors: tuple[DrilldownCell, ...]
    edges: tuple[DrilldownEdge, ...]
    partition: tuple[Scalar, ...] | None = None

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this drilldown."""
        return {
            "series_id": self.series_id,
            "statement_id": self.statement_id,
            "partition": (
                None
                if self.partition is None
                else [jsonable_scalar(value) for value in self.partition]
            ),
            "statements": [stmt.to_dict() for stmt in self.statements],
            "cells": [cell.to_dict() for cell in self.cells],
            "neighbors": [cell.to_dict() for cell in self.neighbors],
            "edges": [edge.to_dict() for edge in self.edges],
        }

    def to_networkx(self):
        """Return a NetworkX MultiDiGraph of drilldown cells and instance edges."""
        return _drilldown_to_networkx(self)


@dataclass(frozen=True, slots=True)
class BundleDrilldown:
    """All classified instance edges that make up one statement-graph bundle."""

    consumer_id: str
    producer_id: str
    access: AccessClass
    distance: int
    coeff: int | None
    offset: int | None
    guarded: bool
    edges: tuple[DrilldownEdge, ...]
    partitions: tuple[tuple[Scalar, ...], ...]

    @property
    def instance_edge_count(self) -> int:
        """Number of canonical instance edges in this bundle."""
        return len(self.edges)

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this bundle expansion."""
        return {
            "consumer_id": self.consumer_id,
            "producer_id": self.producer_id,
            "access": self.access,
            "distance": self.distance,
            "coeff": self.coeff,
            "offset": self.offset,
            "guarded": self.guarded,
            "instance_edge_count": self.instance_edge_count,
            "partitions": [[jsonable_scalar(value) for value in part] for part in self.partitions],
            "edges": [edge.to_dict() for edge in self.edges],
        }


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


def drilldown_partition(
    view: SemanticCatalogView,
    graph: DependencyGraph,
    series_id: str,
    partition: Sequence[Scalar],
) -> SeriesDrilldown:
    """Return cells and adjacency for one instance partition of a series.

    Args:
        view: Statement-partitioned catalog and classified instance edges.
        graph: Cell-level graph used for formula text and values.
        series_id: Bound series whose outer-key block to expand.
        partition: Instance-partition tuple (`schedule.partition_of`).

    Raises:
        SemanticCatalogError: `series_id` is missing, or `partition` matches
            no members of that series.
    """
    series = _require_series(view.catalog, series_id)
    want = tuple(partition)
    selected = tuple(cell for cell in series.cells if view.partition_of(cell) == want)
    if not selected:
        raise SemanticCatalogError(f"unknown partition {want!r} for series {series_id!r}")
    covering = tuple(
        stmt
        for stmt in series.statements
        if any(view.partition_of(cell) == want for cell in stmt.cells)
    )
    return _drilldown(
        view,
        graph,
        series,
        statement=None,
        cells=selected,
        statements=covering,
        partition=want,
        edge_filter="consumer",
    )


def drilldown_bundle(
    view: SemanticCatalogView,
    graph: DependencyGraph,
    *,
    consumer_id: str,
    producer_id: str,
    access: AccessClass,
    distance: int,
    coeff: int | None = None,
    offset: int | None = None,
    guarded: bool = False,
) -> BundleDrilldown:
    """Return every classified instance edge in one statement-graph bundle.

    Args:
        view: Statement-partitioned catalog and classified instance edges.
        graph: Cell-level graph used for formula text and values.
        consumer_id: Statement id of the consuming node.
        producer_id: Statement id of the producing node.
        access: Access class of the bundle.
        distance: Schedule distance of the bundle.
        coeff: Affine coefficient, if the bundle is affine.
        offset: Affine offset, if the bundle is affine.
        guarded: Whether the bundle is a guarded residual.

    Raises:
        SemanticCatalogError: No classified edge matches this bundle key.
    """
    del graph
    matched: list[DrilldownEdge] = []
    partitions: list[tuple[Scalar, ...]] = []
    seen_parts: set[tuple[Scalar, ...]] = set()
    for edge in view.edges.edges:
        consumer_stmt = _statement_id_for(view, edge.consumer_cell)
        producer_stmt = _statement_id_for(view, edge.producer_cell)
        if (
            consumer_stmt != consumer_id
            or producer_stmt != producer_id
            or edge.access != access
            or edge.distance != distance
            or edge.coeff != coeff
            or edge.offset != offset
            or edge.guarded != guarded
        ):
            continue
        record = _as_drilldown_edge(edge, view)
        matched.append(record)
        if record.partition not in seen_parts:
            seen_parts.add(record.partition)
            partitions.append(record.partition)
    if not matched:
        raise SemanticCatalogError(
            f"unknown bundle {consumer_id!r} -> {producer_id!r} "
            f"access={access!r} distance={distance}"
        )
    return BundleDrilldown(
        consumer_id=consumer_id,
        producer_id=producer_id,
        access=access,
        distance=distance,
        coeff=coeff,
        offset=offset,
        guarded=guarded,
        edges=tuple(matched),
        partitions=tuple(partitions),
    )


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


def _statement_id_for(view: SemanticCatalogView, address: CanonicalAddress) -> str:
    covering = _covering_statement(view.catalog, address)
    if covering is not None:
        return covering.statement_id
    owner = view.catalog.series_for(address)
    if owner is not None:
        return owner.series_id
    return REMAINDER_STATEMENT_ID


def _as_drilldown_statement(statement: Statement) -> DrilldownStatement:
    return DrilldownStatement(
        statement_id=statement.statement_id,
        shape_key=statement.shape_key,
        start=statement.start,
        stop=statement.stop,
        cell_count=len(statement.cells),
    )


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
    *,
    edge_filter: str,
) -> tuple[DependenceEdge, ...]:
    candidates = _unique_edges(
        view.edges.by_consumer.get(series_id, ()),
        view.edges.by_producer.get(series_id, ()),
    )
    if selected is None:
        return candidates
    if edge_filter == "consumer":
        return tuple(edge for edge in candidates if edge.consumer_cell in selected)
    return tuple(
        edge
        for edge in candidates
        if edge.consumer_cell in selected or edge.producer_cell in selected
    )


def _as_drilldown_edge(edge: DependenceEdge, view: SemanticCatalogView) -> DrilldownEdge:
    consumer = _covering_statement(view.catalog, edge.consumer_cell)
    producer = _covering_statement(view.catalog, edge.producer_cell)
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
        partition=view.partition_of(edge.consumer_cell),
    )


def _drilldown(
    view: SemanticCatalogView,
    graph: DependencyGraph,
    series: BoundSeries,
    *,
    statement: Statement | None,
    cells: tuple[CanonicalAddress, ...] | None = None,
    statements: tuple[Statement, ...] | None = None,
    partition: tuple[Scalar, ...] | None = None,
    edge_filter: str = "touching",
) -> SeriesDrilldown:
    selected_cells = (
        cells if cells is not None else (statement.cells if statement is not None else series.cells)
    )
    selected = set(selected_cells)
    raw_statements = (
        statements
        if statements is not None
        else ((statement,) if statement is not None else series.statements)
    )
    raw_edges = _series_touching_edges(
        view,
        series.series_id,
        None if statement is None and cells is None else selected,
        edge_filter=edge_filter,
    )
    edges = tuple(_as_drilldown_edge(edge, view) for edge in raw_edges)
    cell_records = tuple(
        _cell_record(
            address,
            view,
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
            view,
            graph,
            drilled_series_id=series.series_id,
        )
        for address in neighbor_addresses
    )
    return SeriesDrilldown(
        series_id=series.series_id,
        statement_id=None if statement is None else statement.statement_id,
        statements=tuple(_as_drilldown_statement(stmt) for stmt in raw_statements),
        cells=cell_records,
        neighbors=neighbors,
        edges=edges,
        partition=partition,
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
            view,
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
    view: SemanticCatalogView,
    graph: DependencyGraph,
    *,
    drilled_series_id: str,
    remainder: bool = False,
) -> DrilldownCell:
    catalog = view.catalog
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
        partition=() if remainder else view.partition_of(address),
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
