"""Graph projection for export and visualization artifacts."""

from __future__ import annotations

import heapq
from collections.abc import Iterable, Iterator, Mapping
from dataclasses import dataclass
from typing import Any, Literal, Protocol

from excel_grapher.core.address_keys import normalize_key
from excel_grapher.grapher.compression import (
    clear_identity_singleton_ref_cache,
    compression_safe_provenance,
    direct_provenance_for_key_in_strings,
    is_identity_transit,
    replace_substrings_at_spans,
)
from excel_grapher.grapher.dependency_provenance import EdgeProvenance
from excel_grapher.grapher.graph import CycleReport, DependencyGraph, NodeKey
from excel_grapher.grapher.guard import GuardExpr
from excel_grapher.grapher.node import Node, NodeView


def copy_dependency_graph(graph: DependencyGraph) -> DependencyGraph:
    """Return a deep copy of `graph` including edges, guards, and provenance."""
    copied = DependencyGraph()
    if graph.sheet_order is not None:
        copied.sheet_order = list(graph.sheet_order)
    if graph.leaf_classification is not None:
        copied.leaf_classification = dict(graph.leaf_classification)
    if graph.named_ranges is not None:
        copied.named_ranges = dict(graph.named_ranges)
    if graph.named_range_ranges is not None:
        copied.named_range_ranges = dict(graph.named_range_ranges)

    for key in graph.keys(order="workbook"):
        node = graph._get_internal_node(key)
        if node is None:
            continue
        copied.add_node(
            Node(
                sheet=node.sheet,
                column=node.column,
                row=node.row,
                formula=node.formula,
                normalized_formula=node.normalized_formula,
                value=node.value,
                is_leaf=node.is_leaf,
                is_target=node.is_target,
                metadata=dict(node.metadata),
            )
        )

    for from_key in graph.keys(order="workbook"):
        for to_key in graph.keys(order="workbook", source=graph.get_dependencies(from_key)):
            attrs = graph.get_edge_attrs(from_key, to_key)
            edge_kwargs: dict[str, Any] = {}
            if attrs.provenance is not None:
                edge_kwargs["provenance"] = attrs.provenance
            copied.add_edge(from_key, to_key, guard=attrs.guard, **edge_kwargs)
    return copied


@dataclass(frozen=True)
class ProjectedNodeSnapshot:
    """Workbook node state captured before projection removes or rewrites it."""

    address: str
    sheet: str
    column: str
    row: int
    formula: str | None
    normalized_formula: str | None
    value: Any
    is_target: bool
    is_leaf: bool
    metadata: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return {
            "address": self.address,
            "sheet": self.sheet,
            "column": self.column,
            "row": self.row,
            "formula": self.formula,
            "normalized_formula": self.normalized_formula,
            "value": self.value,
            "is_target": self.is_target,
            "is_leaf": self.is_leaf,
            "metadata": dict(self.metadata),
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> ProjectedNodeSnapshot:
        return cls(
            address=str(data["address"]),
            sheet=str(data["sheet"]),
            column=str(data["column"]),
            row=int(data["row"]),
            formula=data.get("formula"),
            normalized_formula=data.get("normalized_formula"),
            value=data.get("value"),
            is_target=bool(data.get("is_target", False)),
            is_leaf=bool(data.get("is_leaf", False)),
            metadata=dict(data.get("metadata") or {}),
        )


@dataclass(frozen=True)
class FormulaRewrite:
    """Dependent formula rewrite performed during projection."""

    dependent: str
    before_formula: str | None
    after_formula: str | None
    before_normalized: str | None
    after_normalized: str | None

    def to_dict(self) -> dict[str, Any]:
        return {
            "dependent": self.dependent,
            "before_formula": self.before_formula,
            "after_formula": self.after_formula,
            "before_normalized": self.before_normalized,
            "after_normalized": self.after_normalized,
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> FormulaRewrite:
        return cls(
            dependent=str(data["dependent"]),
            before_formula=data.get("before_formula"),
            after_formula=data.get("after_formula"),
            before_normalized=data.get("before_normalized"),
            after_normalized=data.get("after_normalized"),
        )


@dataclass(frozen=True)
class CollapsedGroup:
    """Collapsed source nodes mapped to one retained computation address."""

    retained: str
    collapsed_sources: tuple[str, ...]
    statement_order: tuple[str, ...]
    external_dependencies: tuple[str, ...]
    node_snapshots: tuple[ProjectedNodeSnapshot, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "retained": self.retained,
            "collapsed_sources": list(self.collapsed_sources),
            "statement_order": list(self.statement_order),
            "external_dependencies": list(self.external_dependencies),
            "node_snapshots": [snap.to_dict() for snap in self.node_snapshots],
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> CollapsedGroup:
        return cls(
            retained=str(data["retained"]),
            collapsed_sources=tuple(str(x) for x in data.get("collapsed_sources") or ()),
            statement_order=tuple(str(x) for x in data.get("statement_order") or ()),
            external_dependencies=tuple(str(x) for x in data.get("external_dependencies") or ()),
            node_snapshots=tuple(
                ProjectedNodeSnapshot.from_dict(item) for item in data.get("node_snapshots") or ()
            ),
        )


@dataclass(frozen=True)
class ProjectionManifest:
    """Serializable lineage for a projected graph artifact."""

    kind: Literal["identity_transit"]
    removed_to_replacement: dict[str, str]
    retained_to_collapsed_sources: dict[str, tuple[str, ...]]
    formula_rewrites: tuple[FormulaRewrite, ...]
    collapsed_groups: tuple[CollapsedGroup, ...]

    @classmethod
    def empty(cls) -> ProjectionManifest:
        """Return an empty identity-transit projection manifest."""
        return cls(
            kind="identity_transit",
            removed_to_replacement={},
            retained_to_collapsed_sources={},
            formula_rewrites=(),
            collapsed_groups=(),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "removed_to_replacement": dict(self.removed_to_replacement),
            "retained_to_collapsed_sources": {
                key: list(value) for key, value in self.retained_to_collapsed_sources.items()
            },
            "formula_rewrites": [item.to_dict() for item in self.formula_rewrites],
            "collapsed_groups": [group.to_dict() for group in self.collapsed_groups],
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> ProjectionManifest:
        retained_raw = data.get("retained_to_collapsed_sources") or {}
        return cls(
            kind="identity_transit",
            removed_to_replacement={
                str(k): str(v) for k, v in (data.get("removed_to_replacement") or {}).items()
            },
            retained_to_collapsed_sources={
                str(k): tuple(str(x) for x in v) for k, v in retained_raw.items()
            },
            formula_rewrites=tuple(
                FormulaRewrite.from_dict(item) for item in data.get("formula_rewrites") or ()
            ),
            collapsed_groups=tuple(
                CollapsedGroup.from_dict(item) for item in data.get("collapsed_groups") or ()
            ),
        )

    def compose(
        self,
        next_manifest: ProjectionManifest,
        *,
        original_graph: DependencyGraph,
    ) -> ProjectionManifest:
        """Compose this manifest with a subsequent projection step.

        `self` maps the original graph to an intermediate graph. `next_manifest`
        maps that intermediate graph to a later projected graph. The composed
        manifest maps original workbook addresses directly to the later projected
        computation addresses.
        """
        if not self.removed_to_replacement:
            return next_manifest
        if not next_manifest.removed_to_replacement:
            return self

        composed_removed: dict[str, str] = {
            removed: next_manifest.map_to_projected(replacement)
            for removed, replacement in self.removed_to_replacement.items()
        }
        for removed, replacement in next_manifest.removed_to_replacement.items():
            composed_removed[removed] = next_manifest.map_to_projected(replacement)

        retained_sources: dict[str, list[str]] = {}
        for removed, replacement in composed_removed.items():
            retained_sources.setdefault(replacement, []).append(removed)
        retained_to_collapsed_sources = {
            retained: tuple(sources) for retained, sources in retained_sources.items()
        }

        snapshots_by_address: dict[str, ProjectedNodeSnapshot] = {}
        for group in self.collapsed_groups + next_manifest.collapsed_groups:
            for snapshot in group.node_snapshots:
                snapshots_by_address[snapshot.address] = snapshot

        collapsed_groups: list[CollapsedGroup] = []
        for retained, sources in retained_to_collapsed_sources.items():
            collapsed_groups.append(
                CollapsedGroup(
                    retained=retained,
                    collapsed_sources=tuple(sources),
                    statement_order=_statement_order(original_graph, sources),
                    external_dependencies=_external_dependencies(
                        original_graph,
                        retained=retained,
                        collapsed_sources=sources,
                    ),
                    node_snapshots=tuple(
                        snapshots_by_address[source]
                        for source in sources
                        if source in snapshots_by_address
                    ),
                )
            )

        return ProjectionManifest(
            kind="identity_transit",
            removed_to_replacement=composed_removed,
            retained_to_collapsed_sources=retained_to_collapsed_sources,
            formula_rewrites=self.formula_rewrites + next_manifest.formula_rewrites,
            collapsed_groups=tuple(collapsed_groups),
        )

    def map_to_projected(self, address: str) -> str:
        """Return the projected computation address for a workbook address."""
        normalized = normalize_key(address)
        return self.removed_to_replacement.get(normalized, normalized)

    def public_aliases_for_export(self, export_addresses: Iterable[str]) -> frozenset[str]:
        """Return public alias addresses whose replacements appear in `export_addresses`."""
        exported = frozenset(normalize_key(addr) for addr in export_addresses)
        aliases: set[str] = set()
        for removed, replacement in self.removed_to_replacement.items():
            if replacement in exported:
                aliases.add(removed)
        return frozenset(aliases)


def _resolve_alias_chain(immediate: Mapping[str, str]) -> dict[str, str]:
    final: dict[str, str] = {}
    for start in immediate:
        current = start
        seen: set[str] = set()
        while current in immediate:
            if current in seen:
                break
            seen.add(current)
            current = immediate[current]
        final[start] = current
    return final


def _node_snapshot(graph: DependencyGraph, key: NodeKey) -> ProjectedNodeSnapshot:
    node = graph._get_internal_node(key)
    if node is None:
        raise KeyError(key)
    return ProjectedNodeSnapshot(
        address=key,
        sheet=node.sheet,
        column=node.column,
        row=node.row,
        formula=node.formula,
        normalized_formula=node.normalized_formula,
        value=node.value,
        is_target=node.is_target,
        is_leaf=node.is_leaf,
        metadata=dict(node.metadata),
    )


def _external_dependencies(
    graph: DependencyGraph,
    *,
    retained: NodeKey,
    collapsed_sources: Iterable[NodeKey],
) -> tuple[str, ...]:
    collapsed = {normalize_key(key) for key in collapsed_sources}
    collapsed.add(normalize_key(retained))
    external: set[str] = set()
    for source in collapsed_sources:
        for dep in graph.get_dependencies(source):
            dep_n = normalize_key(dep)
            if dep_n not in collapsed:
                external.add(dep_n)
    return tuple(sorted(external))


def _statement_order(
    graph: DependencyGraph, collapsed_sources: Iterable[NodeKey]
) -> tuple[str, ...]:
    """Return dependency-first statement order for collapsed source nodes."""
    sources = {normalize_key(key) for key in collapsed_sources}
    if not sources:
        return ()
    try:
        ordered = [key for key in graph.evaluation_order(strict=False) if key in sources]
    except Exception:
        ordered = []
    seen = set(ordered)
    for key in graph.keys(order="workbook", source=sources):
        if key not in seen:
            ordered.append(key)
            seen.add(key)
    return tuple(ordered)


def _compress_one_transit_on_copy(
    graph: DependencyGraph,
    t_key: NodeKey,
    r_key: NodeKey,
    *,
    formula_rewrites: list[FormulaRewrite],
) -> ProjectedNodeSnapshot:
    """Rewrite dependents and remove transit node `t_key` on `graph`."""
    snapshot = _node_snapshot(graph, t_key)
    for d_key in list(graph._reverse_edges.get(t_key, set())):
        extra = graph._edge_extra.get((d_key, t_key), {})
        prov = extra.get("provenance")
        guard = graph._guards.get((d_key, t_key))
        d_node = graph._nodes.get(d_key)
        if d_node is None:
            continue

        before_formula = d_node.formula
        before_normalized = d_node.normalized_formula
        new_formula = before_formula
        new_norm = before_normalized
        if isinstance(prov, EdgeProvenance) and prov.direct_sites_formula and new_formula:
            new_formula = replace_substrings_at_spans(new_formula, prov.direct_sites_formula, r_key)
        elif new_formula and t_key in new_formula:
            new_formula = new_formula.replace(t_key, r_key)

        if isinstance(prov, EdgeProvenance) and prov.direct_sites_normalized and new_norm:
            new_norm = replace_substrings_at_spans(new_norm, prov.direct_sites_normalized, r_key)
        elif new_norm and t_key in new_norm:
            new_norm = new_norm.replace(t_key, r_key)

        if before_formula != new_formula or before_normalized != new_norm:
            formula_rewrites.append(
                FormulaRewrite(
                    dependent=d_key,
                    before_formula=before_formula,
                    after_formula=new_formula,
                    before_normalized=before_normalized,
                    after_normalized=new_norm,
                )
            )

        d_node.formula = new_formula
        d_node.normalized_formula = new_norm

        graph._remove_edge(d_key, t_key)
        new_prov = direct_provenance_for_key_in_strings(new_formula, new_norm, r_key)
        graph.add_edge(d_key, r_key, guard=guard, provenance=new_prov)

    for dep in list(graph._edges.get(t_key, set())):
        graph._remove_edge(t_key, dep)
    graph._nodes.pop(t_key, None)
    graph._edges.pop(t_key, None)
    graph._reverse_edges.pop(t_key, None)
    return snapshot


def _build_identity_transit_manifest(
    original_graph: DependencyGraph,
    *,
    immediate_removed: dict[str, str],
    removal_order: list[str],
    formula_rewrites: list[FormulaRewrite],
    snapshots_by_removed: dict[str, ProjectedNodeSnapshot],
) -> ProjectionManifest:
    removed_to_replacement = _resolve_alias_chain(immediate_removed)
    retained_sources: dict[str, list[str]] = {}
    for removed in removal_order:
        retained = removed_to_replacement[removed]
        retained_sources.setdefault(retained, []).append(removed)

    retained_to_collapsed_sources = {
        retained: tuple(sources) for retained, sources in retained_sources.items() if sources
    }

    collapsed_groups: list[CollapsedGroup] = []
    for retained, sources in retained_to_collapsed_sources.items():
        collapsed_groups.append(
            CollapsedGroup(
                retained=retained,
                collapsed_sources=tuple(sources),
                statement_order=_statement_order(original_graph, sources),
                external_dependencies=_external_dependencies(
                    original_graph,
                    retained=retained,
                    collapsed_sources=sources,
                ),
                node_snapshots=tuple(snapshots_by_removed[source] for source in sources),
            )
        )

    return ProjectionManifest(
        kind="identity_transit",
        removed_to_replacement=removed_to_replacement,
        retained_to_collapsed_sources=retained_to_collapsed_sources,
        formula_rewrites=tuple(formula_rewrites),
        collapsed_groups=tuple(collapsed_groups),
    )


def project_identity_transits(
    graph: DependencyGraph,
) -> tuple[DependencyGraph, ProjectionManifest]:
    """Return a projected copy of `graph` with identity transit nodes collapsed."""
    original_graph = graph
    projected = copy_dependency_graph(graph)
    immediate_removed: dict[str, str] = {}
    removal_order: list[str] = []
    formula_rewrites: list[FormulaRewrite] = []
    snapshots_by_removed: dict[str, ProjectedNodeSnapshot] = {}

    clear_identity_singleton_ref_cache()
    try:
        heap: list[NodeKey] = list(projected._nodes.keys())
        heapq.heapify(heap)
        while heap:
            t_key = heapq.heappop(heap)
            if t_key not in projected._nodes:
                continue
            r_key = is_identity_transit(projected, t_key)
            if r_key is None:
                continue
            dependents_t = projected._reverse_edges.get(t_key, set())
            if not dependents_t:
                continue
            ok = True
            for d_key in dependents_t:
                prov = projected._edge_extra.get((d_key, t_key), {}).get("provenance")
                if not compression_safe_provenance(
                    prov if isinstance(prov, EdgeProvenance) else None
                ):
                    ok = False
                    break
            if not ok:
                continue

            dependents_before = list(dependents_t)
            snapshot = _compress_one_transit_on_copy(
                projected,
                t_key,
                r_key,
                formula_rewrites=formula_rewrites,
            )
            immediate_removed[t_key] = r_key
            removal_order.append(t_key)
            snapshots_by_removed[t_key] = snapshot
            for d_key in dependents_before:
                heapq.heappush(heap, d_key)
    finally:
        clear_identity_singleton_ref_cache()

    manifest = _build_identity_transit_manifest(
        original_graph,
        immediate_removed=immediate_removed,
        removal_order=removal_order,
        formula_rewrites=formula_rewrites,
        snapshots_by_removed=snapshots_by_removed,
    )
    return projected, manifest


@dataclass
class ProjectionResult:
    """Projected graph facade with durable lineage back to the canonical graph."""

    original_graph: DependencyGraph
    projected_graph: DependencyGraph
    manifest: ProjectionManifest

    # ---- graph-like read surface (projected graph) -------------------------

    @property
    def leaf_classification(self) -> dict[str, str] | None:
        return self.projected_graph.leaf_classification

    @leaf_classification.setter
    def leaf_classification(self, value: dict[str, str] | None) -> None:
        self.projected_graph.leaf_classification = value

    @property
    def sheet_order(self) -> list[str] | None:
        return self.projected_graph.sheet_order

    @property
    def named_ranges(self) -> dict[str, tuple[str, str]] | None:
        return self.original_graph.named_ranges

    @property
    def named_range_ranges(self) -> dict[str, tuple[str, str, str]] | None:
        return self.original_graph.named_range_ranges

    def __contains__(self, key: NodeKey) -> bool:
        return key in self.projected_graph

    def __iter__(self) -> Iterator[NodeKey]:
        return iter(self.projected_graph)

    def __len__(self) -> int:
        return len(self.projected_graph)

    def keys(
        self,
        *,
        order: Literal["insertion", "lexical", "workbook"] = "insertion",
        source: Iterable[NodeKey] | None = None,
    ) -> list[NodeKey]:
        return self.projected_graph.keys(order=order, source=source)

    def get_node(self, address: NodeKey) -> NodeView | None:
        return self.projected_graph.get_node(address)

    def get_dependencies(self, address: NodeKey) -> frozenset[NodeKey]:
        return self.projected_graph.get_dependencies(address)

    def get_dependents(self, address: NodeKey) -> frozenset[NodeKey]:
        return self.projected_graph.get_dependents(address)

    def get_edge_attrs(self, from_key: NodeKey, to_key: NodeKey):
        return self.projected_graph.get_edge_attrs(from_key, to_key)

    def get_edge_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None:
        return self.projected_graph.get_edge_guard(from_key, to_key)

    def leaf_keys(self) -> list[NodeKey]:
        return self.projected_graph.leaf_keys()

    def formula_keys(self) -> list[NodeKey]:
        return self.projected_graph.formula_keys()

    def target_keys(self) -> list[NodeKey]:
        """Return target-marked addresses from the canonical workbook graph."""
        return self.original_graph.target_keys()

    def evaluation_order(
        self, *, strict: bool = True, iterate_enabled: bool | None = None
    ) -> list[NodeKey]:
        return self.projected_graph.evaluation_order(
            strict=strict,
            iterate_enabled=iterate_enabled,
        )

    def cycle_report(self) -> CycleReport:
        return self.projected_graph.cycle_report()

    def map_to_projected(self, address: str) -> str:
        """Map a canonical workbook address to its projected computation address."""
        return self.manifest.map_to_projected(address)


class IdentityTransitCompression:
    """Collapse pure identity transit nodes into a projected export graph."""

    def project(self, graph: DependencyGraph) -> ProjectionResult:
        """Build a non-mutating identity-transit projection for export artifacts."""
        projected, manifest = project_identity_transits(graph)
        return ProjectionResult(
            original_graph=graph,
            projected_graph=projected,
            manifest=manifest,
        )


class ProjectionStep(Protocol):
    """Projection step that can build a `ProjectionResult` from a graph."""

    def project(self, graph: DependencyGraph) -> ProjectionResult: ...


def apply_projection(
    graph: DependencyGraph,
    projections: Iterable[ProjectionStep],
) -> ProjectionResult:
    """Apply one or more projection steps, returning the final projection result."""
    current = graph
    projected = copy_dependency_graph(graph)
    manifest = ProjectionManifest.empty()
    applied = False
    for step in projections:
        result = step.project(current)
        manifest = manifest.compose(result.manifest, original_graph=graph)
        current = result.projected_graph
        projected = result.projected_graph
        applied = True
    if not applied:
        return ProjectionResult(
            original_graph=graph,
            projected_graph=projected,
            manifest=manifest,
        )
    return ProjectionResult(
        original_graph=graph,
        projected_graph=projected,
        manifest=manifest,
    )
