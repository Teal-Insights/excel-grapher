"""Graph projection for export and visualization artifacts.

Projections build artifact-specific, non-mutating views of a canonical
`DependencyGraph`. A projection produces a `ProjectionResult` (a read-only graph
facade plus a serializable `ProjectionManifest`) so exporters and visualization
can consume a smaller graph while preserving workbook-facing identity and
lineage back to the original graph.

The manifest separates two concerns:

- `forwarding_map`: value-equivalent removed addresses mapped to retained
  computations. Codegen queries this through `map_to_projected` only for
  addresses that are part of the public export surface.
- collapse lineage (`retained_to_collapsed_sources`, `removed_node_snapshots`,
  `collapsed_groups`, `formula_rewrites`): a durable record of which original
  nodes folded into each retained node, in dependency order, with original
  formulas and metadata. This supports audit and downstream refactoring and does
  not imply value-equivalence.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass
from typing import Any, Literal, Protocol

from excel_grapher.core.address_keys import normalize_key
from excel_grapher.grapher.compression import IdentityTransitCompressionRecord
from excel_grapher.grapher.graph import CycleError, CycleReport, DependencyGraph, NodeKey
from excel_grapher.grapher.guard import GuardExpr
from excel_grapher.grapher.node import NodeView


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
    """Source nodes that folded into one retained computation address.

    Node snapshots are stored once on the manifest's `removed_node_snapshots`
    and referenced here by address via `collapsed_sources` / `statement_order`.
    """

    retained: str
    collapsed_sources: tuple[str, ...]
    statement_order: tuple[str, ...]
    external_dependencies: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "retained": self.retained,
            "collapsed_sources": list(self.collapsed_sources),
            "statement_order": list(self.statement_order),
            "external_dependencies": list(self.external_dependencies),
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> CollapsedGroup:
        return cls(
            retained=str(data["retained"]),
            collapsed_sources=tuple(str(x) for x in data.get("collapsed_sources") or ()),
            statement_order=tuple(str(x) for x in data.get("statement_order") or ()),
            external_dependencies=tuple(str(x) for x in data.get("external_dependencies") or ()),
        )


class ProjectionManifest(Protocol):
    """Stable contract a projection manifest exposes to export and composition.

    Implementations must surface a `kind` tag, an address-mapping helper that
    codegen consumes, plus a serializable `to_dict`. Kind-specific forwarding
    maps and lineage may be carried as additional fields.
    """

    kind: str

    def map_to_projected(self, address: str) -> str: ...

    def to_dict(self) -> dict[str, Any]: ...


def _map_to_projected(forwarding_map: Mapping[str, str], address: str) -> str:
    normalized = normalize_key(address)
    return forwarding_map.get(normalized, normalized)


@dataclass(frozen=True)
class BaseProjectionManifest:
    """General, serializable projection manifest for node-collapsing projections.

    Directly usable by any projection that removes nodes and (optionally) maps
    value-equivalent removed addresses to retained computations; custom projections may also
    subclass to add kind-specific fields. Register `from_dict` under the manifest
    `kind` via `register_projection_manifest` to enable deserialization.
    """

    kind: str
    forwarding_map: dict[str, str]
    retained_to_collapsed_sources: dict[str, tuple[str, ...]]
    removed_node_snapshots: dict[str, ProjectedNodeSnapshot]
    formula_rewrites: tuple[FormulaRewrite, ...]
    collapsed_groups: tuple[CollapsedGroup, ...]

    @classmethod
    def empty(cls, kind: str) -> BaseProjectionManifest:
        """Return an empty manifest tagged with `kind`."""
        return cls(
            kind=kind,
            forwarding_map={},
            retained_to_collapsed_sources={},
            removed_node_snapshots={},
            formula_rewrites=(),
            collapsed_groups=(),
        )

    def map_to_projected(self, address: str) -> str:
        """Return the projected computation address for a workbook address."""
        return _map_to_projected(self.forwarding_map, address)

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "forwarding_map": dict(self.forwarding_map),
            "retained_to_collapsed_sources": {
                key: list(value) for key, value in self.retained_to_collapsed_sources.items()
            },
            "removed_node_snapshots": {
                key: snap.to_dict() for key, snap in self.removed_node_snapshots.items()
            },
            "formula_rewrites": [item.to_dict() for item in self.formula_rewrites],
            "collapsed_groups": [group.to_dict() for group in self.collapsed_groups],
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> BaseProjectionManifest:
        retained_raw = data.get("retained_to_collapsed_sources") or {}
        snapshots_raw = data.get("removed_node_snapshots") or {}
        return cls(
            kind=str(data["kind"]),
            forwarding_map={str(k): str(v) for k, v in (data.get("forwarding_map") or {}).items()},
            retained_to_collapsed_sources={
                str(k): tuple(str(x) for x in v) for k, v in retained_raw.items()
            },
            removed_node_snapshots={
                str(k): ProjectedNodeSnapshot.from_dict(v) for k, v in snapshots_raw.items()
            },
            formula_rewrites=tuple(
                FormulaRewrite.from_dict(item) for item in data.get("formula_rewrites") or ()
            ),
            collapsed_groups=tuple(
                CollapsedGroup.from_dict(item) for item in data.get("collapsed_groups") or ()
            ),
        )


@dataclass(frozen=True)
class CompositeProjectionManifest:
    """Lineage for a sequence of composed projection steps.

    `forwarding_map` is the transitively chained forwarding map across all
    steps (what codegen consumes); `steps` preserves each component manifest for
    audit and serialization. Heterogeneous step kinds are supported.
    """

    forwarding_map: dict[str, str]
    steps: tuple[ProjectionManifest, ...]
    kind: str = "composite"

    def map_to_projected(self, address: str) -> str:
        """Return the projected computation address for a workbook address."""
        projected = normalize_key(address)
        for step in self.steps:
            projected = step.map_to_projected(projected)
        return projected

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "forwarding_map": dict(self.forwarding_map),
            "steps": [step.to_dict() for step in self.steps],
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> CompositeProjectionManifest:
        return cls(
            forwarding_map={str(k): str(v) for k, v in (data.get("forwarding_map") or {}).items()},
            steps=tuple(resolve_projection_manifest(step) for step in data.get("steps") or ()),
        )


ProjectionManifestFromDict = Callable[[Mapping[str, Any]], ProjectionManifest]

_MANIFEST_REGISTRY: dict[str, ProjectionManifestFromDict] = {}


def register_projection_manifest(
    kind: str,
    from_dict: ProjectionManifestFromDict,
    *,
    replace: bool = False,
) -> None:
    """Register a `from_dict` deserializer for projection manifests of `kind`.

    Args:
        kind: Manifest `kind` tag to register.
        from_dict: Callable building a manifest from its serialized mapping.
        replace: Allow overwriting an existing registration.

    Raises:
        ValueError: If `kind` is already registered and `replace` is False.
    """
    if not replace and kind in _MANIFEST_REGISTRY:
        raise ValueError(f"projection manifest kind already registered: {kind!r}")
    _MANIFEST_REGISTRY[kind] = from_dict


def unregister_projection_manifest(kind: str) -> None:
    """Remove a registered projection manifest `kind` (no-op if absent)."""
    _MANIFEST_REGISTRY.pop(kind, None)


def resolve_projection_manifest(data: Mapping[str, Any]) -> ProjectionManifest:
    """Deserialize a manifest by dispatching on its `kind` via the registry.

    Raises:
        ValueError: If the manifest `kind` has no registered deserializer.
    """
    kind = str(data.get("kind", ""))
    from_dict = _MANIFEST_REGISTRY.get(kind)
    if from_dict is None:
        raise ValueError(f"unsupported projection manifest kind: {kind!r}")
    return from_dict(data)


def _projected_snapshot(snapshot: Any) -> ProjectedNodeSnapshot:
    return ProjectedNodeSnapshot(
        address=snapshot.address,
        sheet=snapshot.sheet,
        column=snapshot.column,
        row=snapshot.row,
        formula=snapshot.formula,
        normalized_formula=snapshot.normalized_formula,
        value=snapshot.value,
        is_target=snapshot.is_target,
        is_leaf=snapshot.is_leaf,
        metadata=dict(snapshot.metadata),
    )


def _projected_formula_rewrite(rewrite: Any) -> FormulaRewrite:
    return FormulaRewrite(
        dependent=rewrite.dependent,
        before_formula=rewrite.before_formula,
        after_formula=rewrite.after_formula,
        before_normalized=rewrite.before_normalized,
        after_normalized=rewrite.after_normalized,
    )


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
    except CycleError:
        ordered = []
    seen = set(ordered)
    for key in graph.keys(order="workbook", source=sources):
        if key not in seen:
            ordered.append(key)
            seen.add(key)
    return tuple(ordered)


def build_forwarding_projection_manifest(
    original_graph: DependencyGraph,
    record: IdentityTransitCompressionRecord,
    *,
    kind: str,
) -> BaseProjectionManifest:
    """Build a `BaseProjectionManifest` from a graph compression `record`.

    All removed nodes are treated as value-equivalent forwarding aliases (the
    identity-transit contract). The collapse lineage is keyed back to
    `original_graph` for statement order and external-dependency boundaries.
    """
    forwarding_map = _resolve_alias_chain(record.immediate_removed)
    retained_sources: dict[str, list[str]] = {}
    for removed in record.removal_order:
        retained = forwarding_map[removed]
        retained_sources.setdefault(retained, []).append(removed)

    retained_to_collapsed_sources = {
        retained: tuple(sources) for retained, sources in retained_sources.items() if sources
    }

    removed_node_snapshots = {
        address: _projected_snapshot(snapshot)
        for address, snapshot in record.snapshots_by_removed.items()
    }

    collapsed_groups = tuple(
        CollapsedGroup(
            retained=retained,
            collapsed_sources=tuple(sources),
            statement_order=_statement_order(original_graph, sources),
            external_dependencies=_external_dependencies(
                original_graph,
                retained=retained,
                collapsed_sources=sources,
            ),
        )
        for retained, sources in retained_to_collapsed_sources.items()
    )

    return BaseProjectionManifest(
        kind=kind,
        forwarding_map=forwarding_map,
        retained_to_collapsed_sources=retained_to_collapsed_sources,
        removed_node_snapshots=removed_node_snapshots,
        formula_rewrites=tuple(
            _projected_formula_rewrite(rewrite) for rewrite in record.formula_rewrites
        ),
        collapsed_groups=collapsed_groups,
    )


def project_identity_transits(
    graph: DependencyGraph,
) -> tuple[DependencyGraph, BaseProjectionManifest]:
    """Return a projected copy of `graph` with identity transit nodes collapsed."""
    projected = graph.copy()
    record = IdentityTransitCompressionRecord()
    projected.compress_identity_transits(record=record)
    return projected, build_forwarding_projection_manifest(graph, record, kind="identity_transit")


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
    """Projection step that builds a `ProjectionResult` from a graph."""

    def project(self, graph: DependencyGraph) -> ProjectionResult: ...


def _chain_forwarding_maps(manifests: list[ProjectionManifest]) -> dict[str, str]:
    composed: dict[str, str] = {}
    for index, manifest in enumerate(manifests):
        later = manifests[index + 1 :]
        forwarding_map = getattr(manifest, "forwarding_map", {})
        for removed, replacement in forwarding_map.items():
            final = replacement
            for following in later:
                final = following.map_to_projected(final)
            composed[removed] = final
    return composed


def apply_projection(
    graph: DependencyGraph,
    projections: Iterable[ProjectionStep],
) -> ProjectionResult:
    """Apply one or more projection steps, returning the final projection result.

    Steps are applied in order, each projecting the previous step's projected
    graph. A single step's manifest is returned as-is; multiple steps are folded
    into a `CompositeProjectionManifest` carrying the chained forwarding map and
    every component manifest. Heterogeneous step kinds are supported.
    """
    manifests: list[ProjectionManifest] = []
    projected = graph
    for step in projections:
        result = step.project(projected)
        manifests.append(result.manifest)
        projected = result.projected_graph

    if not manifests:
        return ProjectionResult(
            original_graph=graph,
            projected_graph=graph.copy(),
            manifest=BaseProjectionManifest.empty(kind="empty"),
        )
    if len(manifests) == 1:
        return ProjectionResult(
            original_graph=graph,
            projected_graph=projected,
            manifest=manifests[0],
        )
    return ProjectionResult(
        original_graph=graph,
        projected_graph=projected,
        manifest=CompositeProjectionManifest(
            forwarding_map=_chain_forwarding_maps(manifests),
            steps=tuple(manifests),
        ),
    )


register_projection_manifest("identity_transit", BaseProjectionManifest.from_dict)
register_projection_manifest("empty", BaseProjectionManifest.from_dict)
register_projection_manifest("composite", CompositeProjectionManifest.from_dict)
