"""Statement-level quotient of classified instance edges.

Nodes are statements (one formula-shape run inside a bound series). Edges
are bundles keyed by consumer, producer, access class, schedule distance,
affine coefficients, and guarded status. Instance partitions are metadata
on each bundle, not extra vertices.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical, parse_address
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    schedule_partition,
)
from excel_grapher.exporter.inverted_tree.deps import AccessClass, DependenceEdge
from excel_grapher.exporter.semantic_catalog import SemanticCatalogView
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.lightweight_viz import unconditional_scc_ranks
from excel_grapher.series_bindings.types import Scalar

REMAINDER_STATEMENT_ID = "__unbound__"
MIXED_SHEET = "mixed"

_BundleKey = tuple[
    str,
    str,
    AccessClass,
    int,
    int | None,
    int | None,
    bool,
]


@dataclass(frozen=True, slots=True)
class StatementLabels:
    """Series-level binding documentation copied onto each statement node."""

    notes: str | None = None
    sdmx_notes: str | None = None
    compute_name: str | None = None
    groups: tuple[tuple[str, ...], ...] = ()
    measure_concept: str | None = None
    measure_name: str | None = None
    measure_description: str | None = None
    axis_labels: str | None = None

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of label fields."""
        return {
            "notes": self.notes,
            "sdmx_notes": self.sdmx_notes,
            "compute_name": self.compute_name,
            "groups": [list(path) for path in self.groups],
            "measure_concept": self.measure_concept,
            "measure_name": self.measure_name,
            "measure_description": self.measure_description,
            "axis_labels": self.axis_labels,
        }


EMPTY_STATEMENT_LABELS = StatementLabels()


def _optional_str(value: object) -> str | None:
    return value if isinstance(value, str) else None


def _group_paths(raw: Mapping[str, Any]) -> tuple[tuple[str, ...], ...]:
    refs = raw.get("groups")
    if not isinstance(refs, list):
        return ()
    paths: list[tuple[str, ...]] = []
    for ref in refs:
        if not isinstance(ref, dict):
            continue
        path = ref.get("path")
        if isinstance(path, list) and path:
            paths.append(tuple(str(part) for part in path))
    return tuple(paths)


def _measure_concept_id(raw: Mapping[str, Any]) -> str | None:
    structure = raw.get("structure")
    measure = structure.get("measure") if isinstance(structure, dict) else None
    if not isinstance(measure, dict) or not measure.get("concept"):
        return None
    return str(measure["concept"])


def statement_labels_for(
    series: BoundSeries,
    concepts: Mapping[str, tuple[str | None, str | None]],
) -> StatementLabels:
    """Copy analyst labels from a bound series and its concept-scheme entry."""
    raw = series.raw
    concept_id = _measure_concept_id(raw)
    name, description = concepts.get(concept_id, (None, None)) if concept_id else (None, None)
    return StatementLabels(
        notes=_optional_str(raw.get("notes")),
        sdmx_notes=_optional_str(raw.get("sdmx_notes")),
        compute_name=series.compute_name,
        groups=_group_paths(raw),
        measure_concept=concept_id,
        measure_name=name,
        measure_description=description,
        axis_labels=series.axis_labels,
    )


@dataclass(frozen=True, slots=True)
class StatementNode:
    """One statement (or the unbound remainder) in the compressed graph."""

    statement_id: str
    series_id: str
    shape_key: str | None
    start: int
    stop: int
    cells: tuple[CanonicalAddress, ...]
    cell_count: int
    direction: str
    sheet: str
    is_remainder: bool = False
    labels: StatementLabels = EMPTY_STATEMENT_LABELS


@dataclass(frozen=True, slots=True)
class StatementBundle:
    """Instance edges sharing one statement-level access signature."""

    consumer_id: str
    producer_id: str
    access: AccessClass
    distance: int
    coeff: int | None
    offset: int | None
    guarded: bool
    instance_edge_count: int
    partitions: tuple[tuple[Scalar, ...], ...]
    representative_consumer_cell: CanonicalAddress
    representative_producer_cell: CanonicalAddress


@dataclass(frozen=True, slots=True)
class StatementGraphStats:
    """Reduction from cells and instance edges to statements and bundles."""

    cell_count: int
    statement_count: int
    instance_edge_count: int
    bundle_count: int
    unbound_cell_count: int
    heterogeneous_partition_pair_count: int


@dataclass(frozen=True, slots=True)
class StatementGraph:
    """Compressed statement graph with schedule ranks and 2-D positions."""

    nodes: tuple[StatementNode, ...]
    bundles: tuple[StatementBundle, ...]
    ranks: tuple[int, ...]
    positions: tuple[tuple[float, float], ...]
    stats: StatementGraphStats
    remainder_sample: tuple[CanonicalAddress, ...]
    heterogeneous_pairs: tuple[tuple[str, str], ...]

    def to_networkx(self):
        """Return a NetworkX MultiDiGraph of statements and typed bundles."""
        return statement_graph_to_networkx(self)


def statement_sheet(cells: Sequence[CanonicalAddress]) -> str:
    """Return the shared worksheet name, or `mixed` when sheets disagree.

    Empty cell lists and multi-sheet statements fail closed to `mixed` rather
    than silently taking the first address.
    """
    sheets: set[str] = set()
    for cell in cells:
        sheets.add(parse_address(cell)[0])
        if len(sheets) > 1:
            return MIXED_SHEET
    if len(sheets) != 1:
        return MIXED_SHEET
    return next(iter(sheets))


def _cell_to_statement(catalog: SeriesCatalog) -> dict[CanonicalAddress, str]:
    """Map each bound cell to its covering statement id."""
    mapping: dict[CanonicalAddress, str] = {}
    for series in catalog.series.values():
        for stmt in series.statements:
            for cell in stmt.cells:
                mapping[cell] = stmt.statement_id
        for cell in series.cells:
            mapping.setdefault(cell, series.series_id)
    return mapping


def _statement_nodes(
    catalog: SeriesCatalog,
    concepts: Mapping[str, tuple[str | None, str | None]],
) -> dict[str, StatementNode]:
    """Return statement nodes keyed by statement id, in bindings order."""
    nodes: dict[str, StatementNode] = {}
    for series_id in catalog.order:
        series = catalog.get(series_id)
        labels = statement_labels_for(series, concepts)
        for stmt in series.statements:
            nodes[stmt.statement_id] = StatementNode(
                statement_id=stmt.statement_id,
                series_id=series.series_id,
                shape_key=stmt.shape_key,
                start=stmt.start,
                stop=stmt.stop,
                cells=stmt.cells,
                cell_count=len(stmt.cells),
                direction=series.direction,
                sheet=statement_sheet(stmt.cells),
                labels=labels,
            )
    return nodes


def _ensure_statement_node(
    nodes: dict[str, StatementNode],
    statement_id: str,
    series: BoundSeries,
    labels: StatementLabels,
) -> None:
    if statement_id in nodes:
        return
    nodes[statement_id] = StatementNode(
        statement_id=statement_id,
        series_id=series.series_id,
        shape_key=None,
        start=0,
        stop=len(series.cells),
        cells=series.cells,
        cell_count=len(series.cells),
        direction=series.direction,
        sheet=statement_sheet(series.cells),
        labels=labels,
    )


def _partition_tuple(cell: CanonicalAddress, catalog: SeriesCatalog) -> tuple[Scalar, ...]:
    return schedule_partition(cell, catalog)


def _heterogeneous_pairs(
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> tuple[tuple[str, str], ...]:
    """Return series pairs whose d=0 identity orientation flips by partition.

    Detection is at series grain so a shape-split statement pair does not hide
    the #762 vintage case (opposite residuals in different outer-key blocks).
    """
    by_part: dict[tuple[Scalar, ...], dict[tuple[str, str], bool]] = {}
    for edge in edges:
        if edge.access != "identity" or edge.distance != 0 or edge.guarded:
            continue
        consumer = catalog.series_id_for(edge.consumer_cell)
        producer = catalog.series_id_for(edge.producer_cell)
        if consumer is None or producer is None or consumer == producer:
            continue
        part = _partition_tuple(edge.consumer_cell, catalog)
        directed = by_part.setdefault(part, {})
        directed[(consumer, producer)] = True

    pair_parts: dict[tuple[str, str], dict[str, set[tuple[Scalar, ...]]]] = {}
    for part, directed in by_part.items():
        for consumer, producer in directed:
            key = (consumer, producer) if consumer < producer else (producer, consumer)
            slot = pair_parts.setdefault(key, {"fwd": set(), "rev": set()})
            if (consumer, producer) == key:
                slot["fwd"].add(part)
            else:
                slot["rev"].add(part)

    flagged: list[tuple[str, str]] = []
    for key, slot in pair_parts.items():
        fwd_only = slot["fwd"] - slot["rev"]
        rev_only = slot["rev"] - slot["fwd"]
        if fwd_only and rev_only:
            flagged.append(key)
    flagged.sort()
    return tuple(flagged)


def _accumulate_bundle(
    buckets: dict[_BundleKey, list[DependenceEdge]],
    partitions: dict[_BundleKey, set[tuple[Scalar, ...]]],
    edge: DependenceEdge,
    consumer_id: str,
    producer_id: str,
    catalog: SeriesCatalog,
) -> None:
    key: _BundleKey = (
        consumer_id,
        producer_id,
        edge.access,
        edge.distance,
        edge.coeff,
        edge.offset,
        edge.guarded,
    )
    buckets.setdefault(key, []).append(edge)
    partitions.setdefault(key, set()).add(_partition_tuple(edge.consumer_cell, catalog))


def _bundles_from_edges(
    edges: Sequence[DependenceEdge],
    cell_stmt: Mapping[CanonicalAddress, str],
    catalog: SeriesCatalog,
    remainder_id: str,
) -> tuple[StatementBundle, ...]:
    buckets: dict[_BundleKey, list[DependenceEdge]] = {}
    parts: dict[_BundleKey, set[tuple[Scalar, ...]]] = {}
    for edge in edges:
        consumer_id = cell_stmt.get(edge.consumer_cell, remainder_id)
        producer_id = cell_stmt.get(edge.producer_cell, remainder_id)
        _accumulate_bundle(buckets, parts, edge, consumer_id, producer_id, catalog)

    bundles: list[StatementBundle] = []
    for key in sorted(buckets, key=lambda k: (k[0], k[1], k[2], k[3], k[6], k[4] or 0, k[5] or 0)):
        group = buckets[key]
        first = group[0]
        part_set = tuple(sorted(parts[key], key=lambda p: tuple(str(x) for x in p)))
        bundles.append(
            StatementBundle(
                consumer_id=key[0],
                producer_id=key[1],
                access=key[2],
                distance=key[3],
                coeff=key[4],
                offset=key[5],
                guarded=key[6],
                instance_edge_count=len(group),
                partitions=part_set,
                representative_consumer_cell=first.consumer_cell,
                representative_producer_cell=first.producer_cell,
            )
        )
    return tuple(bundles)


def _rank_and_positions(
    nodes: Sequence[StatementNode],
    bundles: Sequence[StatementBundle],
) -> tuple[tuple[int, ...], tuple[tuple[float, float], ...]]:
    n = len(nodes)
    if n == 0:
        return (), ()
    index = {node.statement_id: i for i, node in enumerate(nodes)}
    adj: list[list[int]] = [[] for _ in range(n)]
    for bundle in bundles:
        if bundle.guarded or bundle.distance != 0 or bundle.access not in {"identity", "affine"}:
            continue
        if bundle.consumer_id == bundle.producer_id:
            continue
        src = index.get(bundle.consumer_id)
        dst = index.get(bundle.producer_id)
        if src is None or dst is None:
            continue
        adj[src].append(dst)
    for row in adj:
        row.sort()
        # Unique targets; Kosaraju accepts duplicates but ranks stay stable.
        if len(row) > 1:
            seen: set[int] = set()
            uniq: list[int] = []
            for v in row:
                if v not in seen:
                    seen.add(v)
                    uniq.append(v)
            row[:] = uniq
    ranks, _scc = unconditional_scc_ranks(adj, n)
    max_r = max(ranks) if ranks else 0
    by_rank: dict[int, list[int]] = {}
    for i, rank in enumerate(ranks):
        by_rank.setdefault(rank, []).append(i)
    positions = [(0.0, 0.0)] * n
    for rank, idxs in sorted(by_rank.items()):
        idxs.sort(key=lambda i: (nodes[i].series_id, nodes[i].start, nodes[i].statement_id))
        denom_y = max(max_r, 0) + 1
        y = 1.0 - 2.0 * (rank + 0.5) / max(denom_y, 1)
        if len(idxs) == 1:
            positions[idxs[0]] = (0.0, y)
            continue
        for j, i in enumerate(idxs):
            x = 2.0 * j / (len(idxs) - 1) - 1.0
            positions[i] = (x, y)
    return tuple(ranks), tuple(positions)


def _unbound_cells(
    graph: DependencyGraph,
    catalog: SeriesCatalog,
    *,
    sample_limit: int = 16,
) -> tuple[tuple[CanonicalAddress, ...], int, str]:
    unbound: list[CanonicalAddress] = []
    for key in graph.keys(order="workbook"):
        if key not in catalog.address_to_id:
            unbound.append(as_canonical(key))
    unbound.sort()
    return tuple(unbound[:sample_limit]), len(unbound), statement_sheet(unbound)


def build_statement_graph(
    view: SemanticCatalogView,
    graph: DependencyGraph,
) -> StatementGraph:
    """Contract instance edges into a statement graph with schedule ranks.

    Args:
        view: Catalog plus classified instance edges.
        graph: Cell-level graph used only for unbound remainder accounting.

    Returns:
        Statement nodes, typed bundles, ranks, and reduction stats.
    """
    catalog = view.catalog
    cell_stmt = _cell_to_statement(catalog)
    series_labels = {
        series_id: statement_labels_for(catalog.get(series_id), view.concepts)
        for series_id in catalog.order
    }
    nodes = _statement_nodes(catalog, view.concepts)
    for edge in view.edges.edges:
        consumer_series = catalog.series_for(edge.consumer_cell)
        producer_series = catalog.series_for(edge.producer_cell)
        if consumer_series is not None:
            _ensure_statement_node(
                nodes,
                cell_stmt.get(edge.consumer_cell, consumer_series.series_id),
                consumer_series,
                series_labels.get(consumer_series.series_id, EMPTY_STATEMENT_LABELS),
            )
        if producer_series is not None:
            _ensure_statement_node(
                nodes,
                cell_stmt.get(edge.producer_cell, producer_series.series_id),
                producer_series,
                series_labels.get(producer_series.series_id, EMPTY_STATEMENT_LABELS),
            )

    remainder_sample, unbound_count, remainder_sheet = _unbound_cells(graph, catalog)
    if unbound_count:
        nodes[REMAINDER_STATEMENT_ID] = StatementNode(
            statement_id=REMAINDER_STATEMENT_ID,
            series_id=REMAINDER_STATEMENT_ID,
            shape_key=None,
            start=0,
            stop=unbound_count,
            cells=remainder_sample,
            cell_count=unbound_count,
            direction="internal",
            sheet=remainder_sheet,
            is_remainder=True,
            labels=EMPTY_STATEMENT_LABELS,
        )

    seen_ids: set[str] = set()
    node_list: list[StatementNode] = []
    for series_id in catalog.order:
        for stmt in catalog.get(series_id).statements:
            if stmt.statement_id in seen_ids or stmt.statement_id not in nodes:
                continue
            seen_ids.add(stmt.statement_id)
            node_list.append(nodes[stmt.statement_id])
    extra = [node for sid, node in nodes.items() if sid not in seen_ids]
    extra.sort(key=lambda n: (not n.is_remainder, n.series_id, n.statement_id))
    node_list.extend(extra)

    remainder_id = REMAINDER_STATEMENT_ID if unbound_count else ""
    bundles = _bundles_from_edges(view.edges.edges, cell_stmt, catalog, remainder_id)
    hetero = _heterogeneous_pairs(view.edges.edges, catalog)
    ranks, positions = _rank_and_positions(node_list, bundles)
    bound_cells = sum(len(series.cells) for series in catalog.series.values())
    stats = StatementGraphStats(
        cell_count=bound_cells,
        statement_count=len(node_list) - (1 if unbound_count else 0),
        instance_edge_count=len(view.edges.edges),
        bundle_count=len(bundles),
        unbound_cell_count=unbound_count,
        heterogeneous_partition_pair_count=len(hetero),
    )
    return StatementGraph(
        nodes=tuple(node_list),
        bundles=bundles,
        ranks=ranks,
        positions=positions,
        stats=stats,
        remainder_sample=remainder_sample,
        heterogeneous_pairs=hetero,
    )


def jsonable_scalar(value: Scalar) -> Any:
    """Return a JSON-serializable form of a key scalar."""
    if isinstance(value, datetime):
        return value.isoformat()
    return value


def _bundle_edge_key(bundle: StatementBundle) -> str:
    return f"{bundle.access}:{bundle.distance}:{bundle.coeff}:{bundle.offset}:{int(bundle.guarded)}"


def statement_graph_to_networkx(graph: StatementGraph):
    """Convert a statement graph to a NetworkX MultiDiGraph.

    Nodes are statement ids. Binding labels and `shape_key` are node
    attributes; formula text is not. Bundles become directed multi-edges
    `consumer -> producer` (depends-on), keyed by access class, distance,
    affine coefficients, and guarded status.

    Raises:
        ImportError: If NetworkX is not installed.
    """
    try:
        import networkx as nx
    except Exception as exc:  # pragma: no cover
        raise ImportError(
            "networkx is not installed; add it to use statement_graph_to_networkx()"
        ) from exc

    nx_graph = nx.MultiDiGraph()
    for index, node in enumerate(graph.nodes):
        attrs: dict[str, Any] = {
            "series_id": node.series_id,
            "shape_key": node.shape_key,
            "start": node.start,
            "stop": node.stop,
            "cell_count": node.cell_count,
            "direction": node.direction,
            "sheet": node.sheet,
            "is_remainder": node.is_remainder,
            "rank": graph.ranks[index],
            "x": graph.positions[index][0],
            "y": graph.positions[index][1],
        }
        attrs.update(node.labels.to_dict())
        nx_graph.add_node(node.statement_id, **attrs)
    for bundle in graph.bundles:
        nx_graph.add_edge(
            bundle.consumer_id,
            bundle.producer_id,
            key=_bundle_edge_key(bundle),
            access=bundle.access,
            distance=bundle.distance,
            coeff=bundle.coeff,
            offset=bundle.offset,
            guarded=bundle.guarded,
            instance_edge_count=bundle.instance_edge_count,
            discharged=bundle.distance > 0 or bundle.access == "shift",
        )
    return nx_graph


__all__ = [
    "EMPTY_STATEMENT_LABELS",
    "MIXED_SHEET",
    "REMAINDER_STATEMENT_ID",
    "StatementBundle",
    "StatementGraph",
    "StatementGraphStats",
    "StatementLabels",
    "StatementNode",
    "build_statement_graph",
    "jsonable_scalar",
    "statement_graph_to_networkx",
    "statement_labels_for",
    "statement_sheet",
]
