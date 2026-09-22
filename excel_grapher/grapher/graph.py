from __future__ import annotations

import array
import copy
import heapq
import warnings
from collections.abc import Callable, Iterable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal, Protocol, SupportsIndex, runtime_checkable

if TYPE_CHECKING:
    from pathlib import Path

    from excel_grapher.core.cell_types import CellType
    from excel_grapher.series_bindings.domains import SeriesDomainIndex

    from .compression import IdentityTransitCompressionRecord, OptimalCompressionRecord
    from .dynamic_refs import DynamicRefConfig
    from .graph_consistency import GraphConsistencyIssue

from excel_grapher.core.address_keys import (
    CellKey,
    normalize_key,
    parse_node_key,
    sort_node_keys,
)
from excel_grapher.core.cell_types import CellTypeEnv
from excel_grapher.core.formula_ast import (
    AstNode,
    bind_axes,
    parse_formula_text,
    parse_preserving_axes_optional,
    rebase_relative_axes,
    replace_resolved_cell_ref,
    retarget_resolved_refs,
    unparse_normalized_formula,
)
from excel_grapher.core.formula_shape import FormulaShapeTable

from .csr_adjacency import CsrCscArrays, build_csr_csc
from .dependency_provenance import DependencyCause, EdgeProvenance, merge_edge_provenance
from .edge_meta import EdgeMetaArrays, EdgeMetaLookup, pack_edge_meta_ids
from .graph_pickle import dumps_graph_blob, loads_graph_blob
from .guard import (
    And,
    CellRef,
    Compare,
    GuardConstraints,
    GuardExpr,
    Not,
    Or,
    intern_guard,
    or_guard,
    rewrite_guard_aliases,
    rewrite_guard_keys,
)
from .may_cycle import identity_alias_map
from .node import Node, NodeKey, NodeView, copy_node, node_to_view

# Sentinel so `set_node_ast(..., formula=None)` can clear the raw audit string
# while omitting `formula=` keeps the existing workbook formula.
_FORMULA_UNSET = object()

NodeHook = Callable[[NodeKey, Node], None]

EdgeKey = tuple[NodeKey, NodeKey]


def _empty_uint32() -> array.array[int]:
    """Return an empty uint32 array for CSR field defaults."""
    return array.array("I")


def _empty_guard_intern() -> list[GuardExpr | None]:
    """Return a 1-based guard intern table whose index `0` slot is unused."""
    return [None]


def _empty_prov_intern() -> list[EdgeProvenance | None]:
    """Return a 1-based provenance intern table whose index `0` slot is unused."""
    return [None]


def _copy_adjacency(adjacency: dict[NodeKey, set[NodeKey]]) -> dict[NodeKey, set[NodeKey]]:
    """Copy neighbor sets, omitting empty entries."""
    return {key: set(neighbors) for key, neighbors in adjacency.items() if neighbors}


def _discard_adjacency(
    adjacency: dict[NodeKey, set[NodeKey]], key: NodeKey, neighbor: NodeKey
) -> None:
    """Remove `neighbor` from `key`'s set; drop the key when the set is empty."""
    neighbors = adjacency.get(key)
    if neighbors is None:
        return
    neighbors.discard(neighbor)
    if not neighbors:
        del adjacency[key]


def _or_merge_optional_guards(parts: list[GuardExpr | None]) -> GuardExpr | None:
    """OR-merge guards; `None` (unconditional) wins, matching `add_edge`."""
    if not parts:
        return None
    merged = parts[0]
    for part in parts[1:]:
        if merged is None or part is None:
            merged = None
        elif merged != part:
            merged = or_guard(merged, part)
    return intern_guard(merged) if merged is not None else None


@dataclass(frozen=True)
class EdgeAttrs:
    """Typed read-only container for dependency-edge attributes.

    Returned by `DependencyGraph.get_edge_attrs`. A missing edge yields an
    `EdgeAttrs` with all fields set to `None`.
    """

    guard: GuardExpr | None = None
    provenance: EdgeProvenance | None = None


@dataclass(frozen=True)
class CycleReport:
    """Result of cycle analysis."""

    has_must_cycles: bool
    has_may_cycles: bool
    must_cycles: list[set[NodeKey]]
    may_cycles: list[set[NodeKey]]
    example_must_cycle_path: list[NodeKey] | None = None
    example_may_cycle_path: list[NodeKey] | None = None


class CycleError(ValueError):
    """Raised when a cycle prevents computing evaluation order."""

    def __init__(self, message: str, cycle_path: list[NodeKey], is_must_cycle: bool):
        super().__init__(message)
        self.cycle_path = cycle_path
        self.is_must_cycle = is_must_cycle


@runtime_checkable
class GraphReadView(Protocol):
    """Read-only dependency-graph surface shared by graphs and projected views.

    Consumers that only read a graph (for example `to_networkx`,
    `to_web_viz_payload`, `CodeGenerator`, and `write_workbook`) can accept any
    object satisfying this protocol, including projected facades such as
    `ProjectionResult`, without depending on the concrete `DependencyGraph`
    type. It captures node iteration, node and edge lookups, key listings,
    leaf/formula/target classification, and evaluation order; mutation is
    intentionally excluded.
    """

    leaf_classification: dict[str, str] | None
    sheet_order: list[str] | None
    named_ranges: dict[str, tuple[str, str]] | None
    named_range_ranges: dict[str, tuple[str, str, str]] | None

    def __contains__(self, key: NodeKey) -> bool: ...

    def __iter__(self) -> Iterator[NodeKey]: ...

    def __len__(self) -> int: ...

    def keys(
        self,
        *,
        order: Literal["insertion", "lexical", "workbook"] = ...,
        source: Iterable[NodeKey] | None = ...,
    ) -> list[NodeKey]: ...

    def get_node(self, address: NodeKey) -> NodeView | None: ...

    def get_dependencies(self, address: NodeKey) -> frozenset[NodeKey]: ...

    def get_dependents(self, address: NodeKey) -> frozenset[NodeKey]: ...

    def resolve_endpoint(self, address: NodeKey) -> NodeKey | None: ...

    def get_dependency_nodes(self, address: NodeKey) -> frozenset[NodeKey]: ...

    def get_edge_attrs(self, from_key: NodeKey, to_key: NodeKey) -> EdgeAttrs: ...

    def get_edge_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None: ...

    def is_guarded(self, from_key: NodeKey, to_key: NodeKey) -> bool: ...

    def leaf_keys(self) -> list[NodeKey]: ...

    def formula_keys(self) -> list[NodeKey]: ...

    def target_keys(self) -> list[NodeKey]: ...

    def evaluation_order(
        self, *, strict: bool = ..., iterate_enabled: bool | None = ..., cell_type_env: Any = ...
    ) -> list[NodeKey]: ...

    def cycle_report(self, *, cell_type_env: Any = ...) -> CycleReport: ...


@dataclass
class DependencyGraph:
    """Mutable workbook dependency graph.

    Node identity is the canonical sheet-qualified cell address (`CellKey`).
    `get_node(key)` looks up that exact stored key.

    `get_dependencies` / `get_dependents` / `get_edge_attrs` return endpoints
    exactly as stored. Topology is uint32 CSR+CSC after `rebuild_adjacency`;
    `_edges` / `_reverse_edges` are a mutation staging buffer and are empty
    once compact. Edge guards and provenance are nnz-parallel intern-id
    arrays beside CSR; `_guards` / `_edge_provenance` are the same kind of
    staging buffer and are empty once compact. Missing adjacency is empty
    (no per-node empty `set()`). `resolve_endpoint` / `get_dependency_nodes`
    resolve to stored cell keys when present (evaluation order, export,
    codegen). Rebuild drops non-node endpoints.

    `formula_shapes` is an optional acceleration overlay from
    `warm_formula_shapes` (unset by default). `Node.formula_ast` is
    authoritative; eval and codegen fall back to AST when the overlay is
    missing. Compression, formula rewrite, and JSON/pickle load drop it
    (`None`). Callers who want shapes again must rewarm; a live
    `FormulaEvaluator` does not pick up a rewarm automatically.
    """

    _nodes: dict[NodeKey, Node] = field(default_factory=dict)
    # Staging adjacency (authoritative while `_staging` is True). Missing key
    # means no neighbors. Emptied after `rebuild_adjacency`.
    _edges: dict[NodeKey, set[NodeKey]] = field(default_factory=dict)
    _reverse_edges: dict[NodeKey, set[NodeKey]] = field(default_factory=dict)
    _staging: bool = field(default=True, repr=False)
    _csr_keys: list[NodeKey] = field(default_factory=list, repr=False)
    _node_index: dict[NodeKey, int] = field(default_factory=dict, repr=False)
    _row_ptr: array.array[int] = field(default_factory=_empty_uint32, repr=False)
    _col_idx: array.array[int] = field(default_factory=_empty_uint32, repr=False)
    _col_ptr: array.array[int] = field(default_factory=_empty_uint32, repr=False)
    _row_idx: array.array[int] = field(default_factory=_empty_uint32, repr=False)
    _guard_id: array.array[int] = field(default_factory=_empty_uint32, repr=False)
    _prov_id: array.array[int] = field(default_factory=_empty_uint32, repr=False)
    _guard_exprs: list[GuardExpr | None] = field(default_factory=_empty_guard_intern, repr=False)
    _provenances: list[EdgeProvenance | None] = field(
        default_factory=_empty_prov_intern, repr=False
    )
    _guards: dict[EdgeKey, GuardExpr] = field(default_factory=dict)
    _edge_provenance: dict[EdgeKey, EdgeProvenance] = field(default_factory=dict)
    _hooks: list[NodeHook] = field(default_factory=list)
    leaf_classification: dict[str, str] | None = None
    sheet_order: list[str] | None = None
    sheet_bounds: dict[str, tuple[int, int]] | None = None
    named_ranges: dict[str, tuple[str, str]] | None = None
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None
    # Opt-in string-keyed AST overlay from warm_ast_cache (not JSON-serialized).
    # Keys: stripped absolute A1 `normalized_formula`. Values: bind_axes trees.
    # Re-warm after load, formula mutation, or move_node that changes targets.
    preparsed_formulas: dict[str, AstNode] | None = None
    # Opt-in eval/codegen overlay from warm_formula_shapes. AST is
    # authoritative; missing overlay falls back. Not JSON/pickle serialized;
    # compression and formula rewrite drop it. Callers must rewarm.
    formula_shapes: FormulaShapeTable | None = None
    # Optional leaf domains used by `cycle_report` when the caller does not pass
    # `cell_type_env`. A `SeriesDomainIndex` is stored as both `domains` and
    # `cell_type_env` so Mapping lookups stay lazy. Pickle/JSON persist a handle
    # (`SeriesDomainHandle`) rather than the expanded table.
    cell_type_env: CellTypeEnv | None = None
    # `(max_branches, max_cells, max_depth)` used when CHOOSE precedents were
    # narrowed during extraction. None on graphs not built by
    # `create_dependency_graph` (consistency then uses default limits).
    dynamic_ref_limits: tuple[int, int, int] | None = None
    domains: SeriesDomainIndex | None = field(default=None, repr=False, compare=False)
    _domains_handle: Any = field(default=None, repr=False, compare=False)
    # Bumped by `set_node_value` so FormulaEvaluator can skip a full leaf poll
    # when no durable value write has happened since the last eager scan.
    _value_generation: int = field(default=0, repr=False, compare=False)

    def copy(self) -> DependencyGraph:
        """Return a deep copy of this graph (node hooks are not copied)."""
        cloned = copy.deepcopy(self)
        cloned._hooks = []
        return cloned

    def __deepcopy__(self, memo: dict[int, Any]) -> DependencyGraph:
        """Clone without going through the pickle blob reduce path."""
        existing = memo.get(id(self))
        if existing is not None:
            return existing
        cloned = self._copy_for_projection()
        memo[id(self)] = cloned
        cloned._hooks = copy.deepcopy(self._hooks, memo)
        return cloned

    def _copy_for_projection(self) -> DependencyGraph:
        """Return an isolated mutable graph clone for projection rewrites."""
        cloned = DependencyGraph()
        cloned._nodes = {key: copy_node(node) for key, node in self._nodes.items()}
        cloned._staging = self._staging
        if self._staging:
            cloned._edges = _copy_adjacency(self._edges)
            cloned._reverse_edges = _copy_adjacency(self._reverse_edges)
            cloned._guards = dict(self._guards)
            cloned._edge_provenance = dict(self._edge_provenance)
        else:
            cloned._edges = {}
            cloned._reverse_edges = {}
            cloned._csr_keys = list(cloned._nodes)
            cloned._node_index = {key: i for i, key in enumerate(cloned._csr_keys)}
            cloned._row_ptr = self._row_ptr[:]
            cloned._col_idx = self._col_idx[:]
            cloned._col_ptr = self._col_ptr[:]
            cloned._row_idx = self._row_idx[:]
            cloned._guard_id = self._guard_id[:]
            cloned._prov_id = self._prov_id[:]
            cloned._guard_exprs = list(self._guard_exprs)
            cloned._provenances = list(self._provenances)
            cloned._guards = {}
            cloned._edge_provenance = {}
        cloned.leaf_classification = (
            dict(self.leaf_classification) if self.leaf_classification is not None else None
        )
        cloned.sheet_order = list(self.sheet_order) if self.sheet_order is not None else None
        cloned.sheet_bounds = dict(self.sheet_bounds) if self.sheet_bounds is not None else None
        cloned.named_ranges = dict(self.named_ranges) if self.named_ranges is not None else None
        cloned.named_range_ranges = (
            dict(self.named_range_ranges) if self.named_range_ranges is not None else None
        )
        cloned.preparsed_formulas = (
            dict(self.preparsed_formulas) if self.preparsed_formulas is not None else None
        )
        cloned.formula_shapes = (
            self.formula_shapes.copy() if self.formula_shapes is not None else None
        )
        if self.domains is not None:
            cloned.domains = self.domains
            cloned.cell_type_env = self.domains
        elif self.cell_type_env is not None and hasattr(self.cell_type_env, "domain_for"):
            cloned.cell_type_env = self.cell_type_env
            cloned.domains = getattr(self, "domains", None)
        else:
            cloned.cell_type_env = (
                dict(self.cell_type_env) if self.cell_type_env is not None else None
            )
            cloned.domains = None
        cloned._domains_handle = self._domains_handle
        cloned.dynamic_ref_limits = self.dynamic_ref_limits
        cloned._value_generation = self._value_generation
        return cloned

    # ---- node insertion and iteration ---------------------------------------

    def add_node(self, node: Node) -> None:
        key = node.key
        is_new = key not in self._nodes
        self._nodes[key] = node
        if is_new and not self._staging:
            self._csr_keys.append(key)
            self._node_index[key] = len(self._csr_keys) - 1
            if not self._row_ptr:
                self._row_ptr.append(0)
                self._col_ptr.append(0)
            self._row_ptr.append(self._row_ptr[-1])
            self._col_ptr.append(self._col_ptr[-1])
        for hook in self._hooks:
            hook(key, node)

    def __contains__(self, key: NodeKey) -> bool:
        return normalize_key(key) in self._nodes

    def __iter__(self) -> Iterator[NodeKey]:
        return iter(self._nodes)

    def __len__(self) -> int:
        return len(self._nodes)

    def keys(
        self,
        *,
        order: Literal["insertion", "lexical", "workbook"] = "insertion",
        source: Iterable[NodeKey] | None = None,
    ) -> list[NodeKey]:
        """Return node keys from `source` (or the graph) using the selected order."""
        key_source: Iterable[NodeKey] = self._nodes if source is None else source
        if order == "insertion":
            return list(key_source)
        if order == "lexical":
            return sorted(key_source)
        if order == "workbook":
            if self.sheet_order:
                return sort_node_keys(key_source, sheet_order=self.sheet_order)
            return sorted(key_source)
        raise ValueError(f"Unsupported key order: {order}")

    # ---- edge insertion -----------------------------------------------------

    def add_edge(
        self,
        from_key: NodeKey,
        to_key: NodeKey,
        *,
        guard: GuardExpr | None = None,
        provenance: EdgeProvenance | None = None,
    ) -> None:
        """Add edge: from_key depends on to_key (from_key -> to_key).

        Endpoints are stored as given (after `normalize_key`). Member-cell
        keys are allowed; callers that need owner nodes should resolve via
        `resolve_endpoint` / `get_dependency_nodes`.

        Re-adding an existing edge merges guards with `or_guard` and merges
        provenance via `merge_edge_provenance`. Omitting `provenance` on a
        re-add leaves any existing provenance unchanged.
        """
        from_key = normalize_key(from_key)
        to_key = normalize_key(to_key)
        ek = (from_key, to_key)
        self._ensure_staging()
        deps_existing = self._edges.get(from_key)
        was_present = deps_existing is not None and to_key in deps_existing

        self._edges.setdefault(from_key, set()).add(to_key)
        self._reverse_edges.setdefault(to_key, set()).add(from_key)

        if not was_present:
            merged_guard = guard
            old_prov: EdgeProvenance | None = None
        else:
            existing_guard = self._guards.get(ek)
            if existing_guard is None or guard is None:
                merged_guard = None
            elif existing_guard == guard:
                merged_guard = guard
            else:
                merged_guard = or_guard(existing_guard, guard)
            old_prov = self._edge_provenance.get(ek)

        if merged_guard is not None:
            self._guards[ek] = intern_guard(merged_guard)
        else:
            self._guards.pop(ek, None)

        if provenance is not None:
            merged_prov = merge_edge_provenance(old_prov, provenance)
            if merged_prov is not None:
                self._edge_provenance[ek] = merged_prov
            else:
                self._edge_provenance.pop(ek, None)
        elif not was_present:
            self._edge_provenance.pop(ek, None)

    # ---- public read API ----------------------------------------------------

    def get_node(self, key: NodeKey) -> NodeView | None:
        """Return an immutable `NodeView` snapshot, or `None` if missing."""
        node = self._nodes.get(normalize_key(key))
        if node is None:
            return None
        return node_to_view(node)

    def get_dependencies(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return an immutable snapshot of `key`'s dependencies (cells it reads).

        Endpoints are returned exactly as stored. Use `get_dependency_nodes`
        when only endpoints that exist as graph nodes are required.
        """
        nk = normalize_key(key)
        if self._staging:
            deps = self._edges.get(nk)
            if not deps:
                return frozenset()
            return frozenset(deps)
        return frozenset(self._csr_dep_keys(nk))

    def get_dependents(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return an immutable snapshot of cells that depend on `key`.

        Endpoints are returned exactly as stored.
        """
        nk = normalize_key(key)
        if self._staging:
            deps = self._reverse_edges.get(nk)
            if not deps:
                return frozenset()
            return frozenset(deps)
        return frozenset(self._csr_dependent_keys(nk))

    def resolve_endpoint(self, key: NodeKey) -> NodeKey | None:
        """Map an edge endpoint to a stored node key when present.

        Returns `None` when the key is not a graph node.
        """
        return self._resolve_graph_endpoint(key)

    def get_dependency_nodes(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return dependencies that exist as stored graph node keys.

        Unresolvable dangling endpoints are omitted.
        """
        out: set[NodeKey] = set()
        for dep in self.get_dependencies(key):
            resolved = self.resolve_endpoint(dep)
            if resolved is not None:
                out.add(resolved)
        return frozenset(out)

    def get_edge_attrs(self, from_key: NodeKey, to_key: NodeKey) -> EdgeAttrs:
        """Return a typed snapshot of the attributes on edge `from_key -> to_key`.

        Lookup uses exact stored endpoints. When the edge does not exist,
        returns an `EdgeAttrs` with all fields set to `None`.
        """
        fk = normalize_key(from_key)
        tk = normalize_key(to_key)
        if self._staging:
            if not self._has_stored_edge(fk, tk):
                return EdgeAttrs()
            return EdgeAttrs(
                guard=self._guards.get((fk, tk)),
                provenance=self._edge_provenance.get((fk, tk)),
            )
        k = self._csr_edge_slot(fk, tk)
        if k is None:
            return EdgeAttrs()
        gid = self._guard_id[k]
        pid = self._prov_id[k]
        return EdgeAttrs(
            guard=self._guard_exprs[gid] if gid else None,
            provenance=self._provenances[pid] if pid else None,
        )

    def get_edge_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None:
        """Return the guard on edge `from_key -> to_key`, or `None` if none.

        Lookup uses exact stored endpoints.
        """
        fk = normalize_key(from_key)
        tk = normalize_key(to_key)
        return self._stored_guard(fk, tk)

    def is_guarded(self, from_key: NodeKey, to_key: NodeKey) -> bool:
        """Return whether edge `from_key -> to_key` carries a guard.

        Membership-only: does not retrieve the guard AST. Prefer this over
        `get_edge_guard(...) is not None` when the expression is unused.
        """
        fk = normalize_key(from_key)
        tk = normalize_key(to_key)
        if self._staging:
            return (fk, tk) in self._guards
        k = self._csr_edge_slot(fk, tk)
        return k is not None and self._guard_id[k] != 0

    def edge_count(self) -> int:
        """Return the number of stored dependency edges."""
        if self._staging:
            return sum(len(deps) for deps in self._edges.values())
        return len(self._col_idx)

    def rebuild_adjacency(self) -> None:
        """Materialize CSR+CSC from current topology and drop dict maps.

        Index space is `_nodes` insertion order. Non-node endpoints are
        dropped. Scratch is `O(n)` uint32; a second Python edge index is not
        built. After rebuild, `_edges`, `_reverse_edges`, `_guards`, and
        `_edge_provenance` are empty. Edge metadata lives in nnz-parallel
        intern-id arrays beside CSR.
        """
        lookup = self._snapshot_edge_meta_lookup()
        nodes = self._nodes
        empty: tuple[NodeKey, ...] = ()
        if self._staging:
            edges = self._edges

            def neighbors_for(key: NodeKey) -> Iterable[NodeKey]:
                return edges.get(key, empty)
        else:

            def neighbors_for(key: NodeKey) -> Iterable[NodeKey]:
                return self._csr_dep_keys(key)

        built = build_csr_csc(tuple(nodes), neighbors_for, is_node=nodes.__contains__)
        meta = pack_edge_meta_ids(built, lookup)
        self._install_csr(built)
        self._install_edge_meta(meta)

    def _install_csr(self, built: CsrCscArrays) -> None:
        self._csr_keys = built.keys
        self._node_index = built.index
        self._row_ptr = built.row_ptr
        self._col_idx = built.col_idx
        self._col_ptr = built.col_ptr
        self._row_idx = built.row_idx
        self._edges = {}
        self._reverse_edges = {}
        self._staging = False

    def _install_edge_meta(self, meta: EdgeMetaArrays) -> None:
        """Install nnz intern-id arrays and drop `EdgeKey` staging maps."""
        self._guard_id = meta.guard_id
        self._prov_id = meta.prov_id
        self._guard_exprs = meta.guard_exprs
        self._provenances = meta.provenances
        self._guards = {}
        self._edge_provenance = {}

    def _ensure_csr(self) -> None:
        if not self._staging:
            return
        self.rebuild_adjacency()

    def _ensure_staging(self) -> None:
        if self._staging:
            return
        self._rehydrate_adjacency_maps()
        self._rehydrate_edge_meta_maps()
        self._clear_csr()
        self._staging = True

    def _clear_csr(self) -> None:
        self._csr_keys = []
        self._node_index = {}
        self._row_ptr = array.array("I")
        self._col_idx = array.array("I")
        self._col_ptr = array.array("I")
        self._row_idx = array.array("I")
        self._guard_id = array.array("I")
        self._prov_id = array.array("I")
        self._guard_exprs = _empty_guard_intern()
        self._provenances = _empty_prov_intern()

    def _rehydrate_adjacency_maps(self) -> None:
        edges: dict[NodeKey, set[NodeKey]] = {}
        reverse: dict[NodeKey, set[NodeKey]] = {}
        keys = self._csr_keys
        row_ptr = self._row_ptr
        col_idx = self._col_idx
        for i, src in enumerate(keys):
            start, end = row_ptr[i], row_ptr[i + 1]
            if start == end:
                continue
            deps: set[NodeKey] = set()
            for k in range(start, end):
                dst = keys[col_idx[k]]
                deps.add(dst)
                reverse.setdefault(dst, set()).add(src)
            edges[src] = deps
        self._edges = edges
        self._reverse_edges = reverse

    def _rehydrate_edge_meta_maps(self) -> None:
        """Rebuild `EdgeKey` staging maps from CSR intern-id arrays."""
        guards: dict[EdgeKey, GuardExpr] = {}
        provenance: dict[EdgeKey, EdgeProvenance] = {}
        keys = self._csr_keys
        row_ptr = self._row_ptr
        col_idx = self._col_idx
        guard_id = self._guard_id
        prov_id = self._prov_id
        guard_exprs = self._guard_exprs
        provenances = self._provenances
        for i, src in enumerate(keys):
            start, end = row_ptr[i], row_ptr[i + 1]
            for k in range(start, end):
                dst = keys[col_idx[k]]
                ek = (src, dst)
                gid = guard_id[k]
                if gid:
                    expr = guard_exprs[gid]
                    if isinstance(expr, GuardExpr):
                        guards[ek] = expr
                pid = prov_id[k]
                if pid:
                    prov = provenances[pid]
                    if isinstance(prov, EdgeProvenance):
                        provenance[ek] = prov
        self._guards = guards
        self._edge_provenance = provenance

    def _snapshot_edge_meta_lookup(self) -> EdgeMetaLookup:
        """Return a `(src, dst)` metadata lookup that outlives rebuild mutation."""
        if self._staging:
            guards = self._guards
            prov = self._edge_provenance

            def lookup(
                src: NodeKey, dst: NodeKey
            ) -> tuple[GuardExpr | None, EdgeProvenance | None]:
                return guards.get((src, dst)), prov.get((src, dst))

            return lookup

        index = self._node_index
        row_ptr = self._row_ptr
        col_idx = self._col_idx
        guard_id = self._guard_id
        prov_id = self._prov_id
        guard_exprs = self._guard_exprs
        provenances = self._provenances

        def lookup(src: NodeKey, dst: NodeKey) -> tuple[GuardExpr | None, EdgeProvenance | None]:
            i = index.get(src)
            j = index.get(dst)
            if i is None or j is None:
                return None, None
            start, end = row_ptr[i], row_ptr[i + 1]
            for k in range(start, end):
                if col_idx[k] == j:
                    gid = guard_id[k]
                    pid = prov_id[k]
                    return (
                        guard_exprs[gid] if gid else None,
                        provenances[pid] if pid else None,
                    )
            return None, None

        return lookup

    def _csr_edge_slot(self, from_key: NodeKey, to_key: NodeKey) -> int | None:
        """Return the CSR nnz index of `from_key -> to_key`, or `None`."""
        i = self._node_index.get(from_key)
        j = self._node_index.get(to_key)
        if i is None or j is None:
            return None
        start, end = self._row_ptr[i], self._row_ptr[i + 1]
        col = self._col_idx
        for k in range(start, end):
            if col[k] == j:
                return k
        return None

    def _stored_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None:
        """Return the stored guard for `from_key -> to_key` in either storage mode."""
        if self._staging:
            v = self._guards.get((from_key, to_key))
            return v if isinstance(v, GuardExpr) else None
        k = self._csr_edge_slot(from_key, to_key)
        if k is None:
            return None
        gid = self._guard_id[k]
        if not gid:
            return None
        expr = self._guard_exprs[gid]
        return expr if isinstance(expr, GuardExpr) else None

    def _stored_provenance(self, from_key: NodeKey, to_key: NodeKey) -> EdgeProvenance | None:
        """Return stored provenance for `from_key -> to_key` in either storage mode."""
        if self._staging:
            return self._edge_provenance.get((from_key, to_key))
        k = self._csr_edge_slot(from_key, to_key)
        if k is None:
            return None
        pid = self._prov_id[k]
        if not pid:
            return None
        prov = self._provenances[pid]
        return prov if isinstance(prov, EdgeProvenance) else None

    def _iter_guard_exprs(self) -> Iterator[GuardExpr]:
        """Iterate interned guard trees (staging map values or compact intern table)."""
        if self._staging:
            yield from self._guards.values()
            return
        for expr in self._guard_exprs:
            if isinstance(expr, GuardExpr):
                yield expr

    def _csr_dep_keys(self, key: NodeKey) -> tuple[NodeKey, ...]:
        i = self._node_index.get(key)
        if i is None:
            return ()
        start, end = self._row_ptr[i], self._row_ptr[i + 1]
        if start == end:
            return ()
        keys = self._csr_keys
        col = self._col_idx
        return tuple(keys[col[k]] for k in range(start, end))

    def _csr_dependent_keys(self, key: NodeKey) -> tuple[NodeKey, ...]:
        j = self._node_index.get(key)
        if j is None:
            return ()
        start, end = self._col_ptr[j], self._col_ptr[j + 1]
        if start == end:
            return ()
        keys = self._csr_keys
        rows = self._row_idx
        return tuple(keys[rows[k]] for k in range(start, end))

    def _iter_dep_keys(self, key: NodeKey) -> Iterator[NodeKey]:
        if self._staging:
            yield from self._edges.get(key, ())
            return
        yield from self._csr_dep_keys(key)

    def _iter_dependent_keys(self, key: NodeKey) -> Iterator[NodeKey]:
        if self._staging:
            yield from self._reverse_edges.get(key, ())
            return
        yield from self._csr_dependent_keys(key)

    def _iter_unguarded_dep_keys(self, key: NodeKey) -> Iterator[NodeKey]:
        if self._staging:
            for dep in self._edges.get(key, ()):
                if (key, dep) in self._guards:
                    continue
                resolved = self._resolve_graph_endpoint(dep)
                if resolved is not None:
                    yield resolved
            return
        i = self._node_index.get(key)
        if i is None:
            return
        keys = self._csr_keys
        col = self._col_idx
        guard_id = self._guard_id
        start, end = self._row_ptr[i], self._row_ptr[i + 1]
        for k in range(start, end):
            if guard_id[k]:
                continue
            yield keys[col[k]]

    def _csr_neighbor_ids(self, i: int, *, unguarded_only: bool = False) -> Iterator[int]:
        start, end = self._row_ptr[i], self._row_ptr[i + 1]
        col = self._col_idx
        if not unguarded_only:
            for k in range(start, end):
                yield col[k]
            return
        guard_id = self._guard_id
        for k in range(start, end):
            if not guard_id[k]:
                yield col[k]

    def _has_stored_edge(self, from_key: NodeKey, to_key: NodeKey) -> bool:
        if self._staging:
            deps = self._edges.get(from_key)
            return deps is not None and to_key in deps
        return self._csr_edge_slot(from_key, to_key) is not None

    # ---- durable node mutation ---------------------------------------------

    def set_node_value(self, key: NodeKey, value: Any) -> None:
        """Set a node's `value` field durably. Raises `KeyError` if missing.

        Increments `_value_generation` so evaluators can skip polling every
        leaf when no durable write has occurred since the last scan.
        """
        nk = normalize_key(key)
        node = self._nodes.get(nk)
        if node is None:
            raise KeyError(f"Cell {key} not found in graph")
        node.value = value
        self._value_generation += 1

    def set_node_metadata(self, key: NodeKey, metadata: Mapping[str, Any]) -> None:
        """Replace a node's metadata mapping durably.

        The provided mapping is copied; subsequent mutations to the caller's
        object do not affect graph state. Raises `KeyError` if the node is
        missing.
        """
        nk = normalize_key(key)
        node = self._nodes.get(nk)
        if node is None:
            raise KeyError(f"Cell {key} not found in graph")
        node.set_metadata(metadata)

    def set_node_formula(
        self,
        key: NodeKey,
        formula: str | None,
        normalized_formula: str | None,
    ) -> None:
        """Set a node's `formula` and parse it into `formula_ast`.

        Parses raw `formula` with axis intent when present; otherwise parses
        `normalized_formula` with the host cell as anchor when available
        (absolute-only when it is not). A successful parse is the source of
        the derived `normalized_formula` view. Unparseable formulas leave
        `formula_ast` unset and keep `normalized_formula` as fallback text.
        Edges are not recomputed; callers rewiring dependencies must update
        edges explicitly. Intended for projection authors building export-only
        graph views. For a topology-aware durable edit, use
        `replace_node_formula`. Drops `formula_shapes`; callers who want the
        overlay must rewarm. Does not validate formula/edge agreement; call
        `validate_consistency` after rewiring.

        Raises:
            KeyError: If the node is missing.
        """
        nk = normalize_key(key)
        node = self._nodes.get(nk)
        if node is None:
            raise KeyError(f"Cell {key} not found in graph")
        node.formula = formula
        ast: AstNode | None = None
        if formula is not None:
            ast = parse_preserving_axes_optional(formula, anchor=node.address or nk)
        if ast is None:
            ast = parse_formula_text(normalized_formula, anchor=node.address or nk)
        node.formula_ast = ast
        if ast is not None:
            node._unparseable_formula = None
        else:
            node._unparseable_formula = normalized_formula
        self._invalidate_formula_shapes()

    def replace_node_formula(
        self,
        key: NodeKey,
        formula: str | None,
        normalized_formula: str | None,
        *,
        workbook: str | Path | None = None,
        dynamic_refs: DynamicRefConfig | None = None,
        use_cached_dynamic_refs: bool = False,
        load_values: bool = True,
        capture_dependency_provenance: bool = True,
        max_depth: int = 50,
        expand_ranges: bool = True,
        max_range_cells: int | None = None,
    ) -> None:
        """Replace a node's formula and rewire outgoing edges from extraction.

        Unlike `set_node_formula` (the projection primitive), this recomputes
        leaf/formula state, outgoing guards, and provenance from the new
        formula. Newly referenced cells are materialized when `workbook` is
        provided. Named ranges that are not on the graph, dynamic
        OFFSET/INDIRECT/INDEX that need the book, and missing subgraph cells
        fail closed with `WorkbookContextRequiredError` when `workbook` is
        omitted — the graph is left unchanged.

        Incoming edges (dependents of this cell) are preserved. `formula_shapes`
        is dropped; callers who want the overlay must rewarm.

        Args:
            key: Existing node to edit.
            formula: New raw formula text, or `None` to clear the formula and
                become a leaf.
            normalized_formula: Fallback parse text when `formula` is
                unparseable. Ignored when `formula` parses.
            workbook: Source `.xlsx` used to resolve named ranges, dynamic
                refs, and to materialize off-path cells.
            dynamic_refs: Constraint-based dynamic-ref config forwarded to
                extraction when `workbook` is set.
            use_cached_dynamic_refs: Resolve OFFSET/INDIRECT/INDEX from cached
                workbook values (requires `workbook`).
            load_values: Load cached Excel values for newly materialized cells.
            capture_dependency_provenance: Record extraction provenance on new
                edges (default True).
            max_depth: BFS depth when materializing from `workbook`.
            expand_ranges: Expand rectangular refs to member cells.
            max_range_cells: Expansion budget; defaults to
                `DEFAULT_MAX_RANGE_CELLS`.

        Raises:
            KeyError: If the node is missing.
            WorkbookContextRequiredError: If workbook context is required and
                `workbook` is omitted.
        """
        from .formula_replace import replace_node_formula as _replace
        from .parser import DEFAULT_MAX_RANGE_CELLS as _DEFAULT_MAX_RANGE_CELLS

        _replace(
            self,
            key,
            formula,
            normalized_formula,
            workbook=workbook,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=use_cached_dynamic_refs,
            load_values=load_values,
            capture_dependency_provenance=capture_dependency_provenance,
            max_depth=max_depth,
            expand_ranges=expand_ranges,
            max_range_cells=(
                _DEFAULT_MAX_RANGE_CELLS if max_range_cells is None else max_range_cells
            ),
        )

    def set_node_ast(
        self,
        key: NodeKey,
        formula_ast: AstNode | None,
        *,
        formula: str | None | object = _FORMULA_UNSET,
    ) -> None:
        """Set a node's formula from `formula_ast`.

        Derives `normalized_formula` as absolute A1. Omit `formula` to leave
        the raw audit string unchanged. Pass `formula=` to set or clear it.
        Unset `formula_ast` clears the derived formula view and, unless
        `formula=` is passed, the raw audit string. Edges are not recomputed;
        callers rewiring dependencies must update edges explicitly. Drops
        `formula_shapes`; callers who want the overlay must rewarm. Does not
        validate formula/edge agreement; call `validate_consistency` after
        rewiring.

        Raises:
            KeyError: If the node is missing.
        """
        nk = normalize_key(key)
        node = self._nodes.get(nk)
        if node is None:
            raise KeyError(f"Cell {key} not found in graph")
        if formula is not _FORMULA_UNSET:
            node.formula = formula if isinstance(formula, str) or formula is None else None
        elif formula_ast is None:
            node.formula = None
        node.formula_ast = formula_ast
        node._unparseable_formula = None
        self._invalidate_formula_shapes()

    def remove_node(self, key: NodeKey) -> None:
        """Remove a node and all of its incident edges.

        Both outgoing dependency edges and incoming dependent edges are dropped,
        along with their guards and provenance. Dependent formulas are not
        rewritten; callers collapsing nodes must update dependents explicitly.
        Node hooks are not invoked. Absent keys are a no-op. Does not
        validate remaining formulas; call `validate_consistency` after
        collapsing nodes.
        """
        nk = normalize_key(key)
        if nk not in self._nodes:
            return
        self._ensure_staging()
        for dep in list(self._edges.get(nk, set())):
            self._remove_edge(nk, dep)
        for dependent in list(self._reverse_edges.get(nk, set())):
            self._remove_edge(dependent, nk)
        self._nodes.pop(nk, None)
        self._edges.pop(nk, None)
        self._reverse_edges.pop(nk, None)
        self.rebuild_adjacency()

    def consistency_issues(self) -> tuple[GraphConsistencyIssue, ...]:
        """Return structured formula/edge/flag disagreements (empty if consistent)."""
        from .graph_consistency import collect_graph_consistency_issues

        return collect_graph_consistency_issues(self)

    def validate_consistency(self) -> None:
        """Raise `GraphConsistencyError` if formulas, edges, or flags disagree.

        Does not rewrite the graph. Opt-in for projection authors after
        `set_node_formula` / `add_edge` / `remove_node`.

        Raises:
            GraphConsistencyError: If any structured consistency issue is found.
        """
        from .graph_consistency import validate_graph_consistency

        validate_graph_consistency(self)

    def move_node(self, old_key: NodeKey, new_key: NodeKey) -> None:
        """Move a node to a new cell, preserving resolved formula targets.

        Excel cut/move: relative axes on the moved host are rewritten so they
        still resolve to the same cells against `new_key`. Dependents that
        resolved to `old_key` are rewritten to resolve to `new_key`. Absolute
        axes keep their kind. `normalized_formula` is a derived absolute-A1
        view and follows the rewritten AST. When `formula_shapes` is warmed,
        affected bindings are re-interned (shared relative params may split or
        join).

        Incoming edges from occupancy (range expansion, whole-column/row
        members) are dropped when the dependent formula was not rewritten to
        mention `new_key`. Remaining `GuardExpr` trees rewrite `CellRef` keys
        and `RangeRef` endpoints. If the destination already has an incoming
        edge from the same dependent, guards merge with `or_guard` (`None`
        unconditional wins) and provenance merges with `merge_edge_provenance`.

        Assigning `Node.address` directly raises; geometry edits must go
        through this method.

        Args:
            old_key: Existing node key.
            new_key: Destination single-cell key.

        Raises:
            KeyError: If `old_key` is missing.
            ValueError: If `new_key` is not a single cell, already exists,
                an affected formula is unparseable, or a range guard would
                split across sheets.
        """
        old_nk = normalize_key(old_key)
        dest = parse_node_key(normalize_key(new_key))
        if not isinstance(dest, CellKey):
            raise ValueError(f"move_node destination must be a single cell, got {new_key!r}")
        new_nk = str(dest)
        if old_nk == new_nk:
            return
        node = self._nodes.get(old_nk)
        if node is None:
            raise KeyError(f"Cell {old_key} not found in graph")
        if new_nk in self._nodes:
            raise ValueError(f"Cell {new_key} already exists in graph")

        dependents = [key for key in self._iter_dependent_keys(old_nk) if key != old_nk]
        self._require_rewritable_formula(node, old_nk)
        for dep_key in dependents:
            dep_node = self._nodes.get(dep_key)
            if dep_node is None:
                continue
            self._require_rewritable_formula(dep_node, dep_key)
        for guard in self._iter_guard_exprs():
            rewrite_guard_keys(guard, old_nk, new_nk)

        if node.formula_ast is not None:
            rebased = rebase_relative_axes(node.formula_ast, old_anchor=old_nk, new_anchor=new_nk)
            node.formula_ast = retarget_resolved_refs(
                rebased, old_key=old_nk, new_key=new_nk, anchor=new_nk
            )

        rewritten_dependents: list[NodeKey] = []
        for dep_key in dependents:
            dep_node = self._nodes.get(dep_key)
            if dep_node is None or dep_node.formula_ast is None:
                continue
            new_ast = retarget_resolved_refs(
                dep_node.formula_ast,
                old_key=old_nk,
                new_key=new_nk,
                anchor=dep_node.address or dep_key,
            )
            if new_ast is dep_node.formula_ast:
                continue
            dep_node.formula_ast = new_ast
            rewritten_dependents.append(dep_key)

        self._relocate_graph_identity(
            old_nk, new_nk, node, rewritten_dependents=frozenset(rewritten_dependents)
        )
        self._refresh_move_provenance(new_nk, rewritten_dependents)
        self._reintern_moved_formula_shapes(old_nk, new_nk, rewritten_dependents)
        self.rebuild_adjacency()

    def _require_rewritable_formula(self, node: Node, key: NodeKey) -> None:
        if node.formula_ast is None and node.has_formula:
            raise ValueError(f"cannot rewrite unparseable formula at {key} when moving a node")

    def _relocate_graph_identity(
        self,
        old_key: NodeKey,
        new_key: NodeKey,
        node: Node,
        rewritten_dependents: frozenset[NodeKey],
    ) -> None:
        self._ensure_staging()
        old_edges = {src: set(dsts) for src, dsts in self._edges.items() if dsts}
        old_guards = dict(self._guards)
        old_prov = dict(self._edge_provenance)

        self._nodes.pop(old_key)
        node._relocate(new_key)
        self._nodes[new_key] = node

        def swap(key: NodeKey) -> NodeKey:
            return new_key if key == old_key else key

        def keep_incoming(src: NodeKey, dst: NodeKey) -> bool:
            if dst != old_key:
                return True
            return src == old_key or src in rewritten_dependents

        guard_parts: dict[EdgeKey, list[GuardExpr | None]] = {}
        prov_parts: dict[EdgeKey, list[EdgeProvenance | None]] = {}
        new_edges: dict[NodeKey, set[NodeKey]] = {}
        new_reverse: dict[NodeKey, set[NodeKey]] = {}
        for src, dests in old_edges.items():
            for dst in sorted(dests):
                if not keep_incoming(src, dst):
                    continue
                new_src, new_dst = swap(src), swap(dst)
                new_edges.setdefault(new_src, set()).add(new_dst)
                new_reverse.setdefault(new_dst, set()).add(new_src)
                new_ek = (new_src, new_dst)
                guard_parts.setdefault(new_ek, []).append(old_guards.get((src, dst)))
                prov_parts.setdefault(new_ek, []).append(old_prov.get((src, dst)))

        self._edges = new_edges
        self._reverse_edges = new_reverse

        self._guards = {}
        for ek, parts in guard_parts.items():
            merged = _or_merge_optional_guards(parts)
            if merged is not None:
                self._guards[ek] = rewrite_guard_keys(merged, old_key, new_key)

        self._edge_provenance = {}
        for ek, parts in prov_parts.items():
            merged_prov = None
            for part in parts:
                merged_prov = merge_edge_provenance(merged_prov, part)
            if merged_prov is not None:
                self._edge_provenance[ek] = merged_prov

        if self.leaf_classification is not None and old_key in self.leaf_classification:
            self.leaf_classification[new_key] = self.leaf_classification.pop(old_key)

    def _refresh_move_provenance(self, moved_key: NodeKey, dependents: list[NodeKey]) -> None:
        from .compression import refresh_direct_sites

        keys = [moved_key, *dependents]
        for key in keys:
            node = self._nodes.get(key)
            if node is None:
                continue
            normalized = node.normalized_formula
            for dep in self._iter_dep_keys(key):
                prov = self._edge_provenance.get((key, dep))
                if prov is None or DependencyCause.direct_ref not in prov.causes:
                    continue
                self._edge_provenance[(key, dep)] = refresh_direct_sites(
                    prov, new_normalized=normalized, precedent_key=dep
                )

    def _reintern_moved_formula_shapes(
        self,
        old_key: NodeKey,
        new_key: NodeKey,
        dependents: list[NodeKey],
    ) -> None:
        table = self.formula_shapes
        if table is None:
            return
        table.drop_binding(old_key)
        for key in (new_key, *dependents):
            node = self._nodes.get(key)
            if node is not None and node.formula_ast is not None:
                table.rebind(key, node.formula_ast)
            else:
                table.drop_binding(key)
        table.prune_unused_shapes()

    # ---- internal accessors -------------------------------------------------

    def _get_internal_node(self, key: NodeKey) -> Node | None:
        """Internal accessor for the live stored `Node` (normalizes key).

        Internal-only: external callers must use `get_node` which returns an
        immutable `NodeView`.
        """
        return self._nodes.get(normalize_key(key))

    # ---- hooks --------------------------------------------------------------

    def register_hook(self, hook: NodeHook) -> None:
        self._hooks.append(hook)

    # ---- classifications / iterators ---------------------------------------

    def leaves(self) -> Iterator[NodeKey]:
        """Iterate over keys of leaf nodes (no dependencies)."""
        for key, node in self._nodes.items():
            if node.is_leaf:
                yield key

    def formula_nodes(self) -> Iterator[tuple[NodeKey, Node]]:
        """Iterate over (key, node) pairs for nodes that contain formulas.

        Formula cells are identified by `has_formula` (`formula_ast` or
        unparseable formula text). `normalized_formula` is a derived view.
        """
        for key, node in self._nodes.items():
            if node.has_formula:
                yield key, node

    def leaf_node_items(self) -> Iterator[tuple[NodeKey, Node]]:
        """Iterate over (key, node) pairs for leaf nodes (no cell dependencies)."""
        for key, node in self._nodes.items():
            if node.is_leaf:
                yield key, node

    def formula_keys(self) -> list[NodeKey]:
        """Return sorted list of keys for nodes that contain formulas."""
        return self.keys(
            order="workbook",
            source=(k for k, node in self._nodes.items() if node.has_formula),
        )

    def leaf_keys(self) -> list[NodeKey]:
        """Return sorted list of keys for nodes with no dependency edges (leaves)."""
        return self.keys(
            order="workbook", source=(k for k, node in self._nodes.items() if node.is_leaf)
        )

    def domain_for(self, key: NodeKey) -> CellType | None:
        """Return the bindings-owned domain for `key`, if one is attached."""
        if self.domains is not None:
            return self.domains.domain_for(str(key))
        if self.cell_type_env is None:
            return None
        getter = getattr(self.cell_type_env, "get", None)
        if getter is None:
            return None
        return getter(str(key))

    def attach_domains(
        self,
        bindings: Mapping[str, Any],
        *,
        workbook: Path | str | None = None,
        bindings_path: Path | str | None = None,
    ) -> SeriesDomainIndex:
        """Attach a lazy series-domain index without rebuilding the graph.

        Args:
            bindings: Loaded (merged) series binding manifest.
            workbook: Workbook path for range expansion and `from_workbook` reads.
            bindings_path: Sidecar path stored on the pickle/JSON handle.

        Returns:
            The attached `SeriesDomainIndex` (also exposed as `cell_type_env`).
        """
        from excel_grapher.series_bindings.domains import SeriesDomainIndex

        index = SeriesDomainIndex.from_bindings(
            bindings, workbook=workbook, graph=self, bindings_path=bindings_path
        )
        self.domains = index
        self.cell_type_env = index
        self._domains_handle = index.handle
        return index

    def target_keys(self) -> list[NodeKey]:
        """Return sorted list of keys marked as original build targets."""
        return self.keys(
            order="workbook", source=(k for k, node in self._nodes.items() if node.is_target)
        )

    def roots(self) -> Iterator[NodeKey]:
        """Iterate keys with no in-graph dependents."""
        for key in self._nodes:
            if not any(self._iter_dependent_keys(key)):
                yield key

    # ---- adjacency helpers for cycle/order analysis ------------------------

    def _resolve_graph_endpoint(self, key: NodeKey) -> NodeKey | None:
        """Map an edge endpoint to a stored node key when present."""
        nk = normalize_key(key)
        if nk in self._nodes:
            return nk
        return None

    def cycle_report(self, *, cell_type_env: CellTypeEnv | None = None) -> CycleReport:
        """Classify must-cycles vs may-cycles, dropping guard-infeasible SCCs.

        When `cell_type_env` is omitted, uses `self.cell_type_env` (set by
        `create_dependency_graph` from `DynamicRefConfig`). Identity-formula
        cells are rewritten to their copied cell before guards are conjoined, so
        singleton leaf domains apply through aliases.
        """
        env = self.cell_type_env if cell_type_env is None else cell_type_env
        aliases = identity_alias_map(self._nodes)
        self._ensure_csr()
        n = len(self._csr_keys)
        keys = self._csr_keys

        def all_neighbors(i: int) -> Iterator[int]:
            return self._csr_neighbor_ids(i)

        def uncond_neighbors(i: int) -> Iterator[int]:
            return self._csr_neighbor_ids(i, unguarded_only=True)

        must_id_sccs = _scc_cycle_ids(n, uncond_neighbors)
        must_sccs = [{keys[i] for i in scc} for scc in must_id_sccs]
        example_must = _find_cycle_path_ids(uncond_neighbors, keys, must_id_sccs)

        may_sccs: list[set[NodeKey]] = []
        example_may: list[NodeKey] | None = None
        for scc_ids in _scc_cycle_ids(n, all_neighbors):
            scc = {keys[i] for i in scc_ids}
            if _subgraph_has_cycle_ids(set(scc_ids), uncond_neighbors):
                continue
            if not _subgraph_has_feasible_cycle(self, scc, cell_type_env=env, aliases=aliases):
                continue
            may_sccs.append(scc)

        if may_sccs:
            example_may = _find_feasible_cycle_path(
                self, may_sccs[0], cell_type_env=env, aliases=aliases
            )

        return CycleReport(
            has_must_cycles=bool(must_sccs),
            has_may_cycles=bool(may_sccs),
            must_cycles=must_sccs,
            may_cycles=may_sccs,
            example_must_cycle_path=example_must,
            example_may_cycle_path=example_may,
        )

    def _workbook_sorted_keys(self, keys: Iterable[NodeKey]) -> list[NodeKey]:
        """Return `keys` sorted by workbook sheet order, then row, then column."""
        materialized = list(keys)
        if not materialized:
            return []
        if self.sheet_order:
            return sort_node_keys(materialized, sheet_order=self.sheet_order)
        return sorted(materialized)

    def evaluation_order(
        self,
        *,
        strict: bool = True,
        iterate_enabled: bool | None = None,
        cell_type_env: CellTypeEnv | None = None,
    ) -> list[NodeKey]:
        """Return nodes in dependency-first order (leaves before formulas that use them).

        Edge direction is A -> B meaning A depends on B. This method returns an
        ordering suitable for sequential evaluation (dependencies first).

        If `iterate_enabled` is True (workbook has iterative calculation on), any
        must-cycle or may-cycle is rejected: generated Python does not emulate Excel's
        iterative convergence. Pass `False` or `None` to apply the usual strict /
        non-strict rules without this check.

        `cell_type_env` is forwarded to `cycle_report` so leaf-domain constraints
        can prove guarded cycles infeasible.
        """
        report = self.cycle_report(cell_type_env=cell_type_env)
        if iterate_enabled is True:
            if report.has_must_cycles:
                raise CycleError(
                    "Iterative calculation is enabled in the workbook, but unconditional "
                    "dependency cycles cannot be reproduced in generated code; break the cycle "
                    "or set calcPr iterate to 0 in the workbook, which may change Excel results.",
                    report.example_must_cycle_path or [],
                    is_must_cycle=True,
                )
            if report.has_may_cycles:
                raise CycleError(
                    "Iterative calculation is enabled in the workbook, but guarded (may-) "
                    "dependency cycles cannot be reproduced in generated code; break the cycle "
                    "or set calcPr iterate to 0 in the workbook, which may change Excel results.",
                    report.example_may_cycle_path or [],
                    is_must_cycle=False,
                )
        if report.has_must_cycles:
            raise CycleError(
                "Must-cycle detected; cannot compute evaluation order",
                report.example_must_cycle_path or [],
                is_must_cycle=True,
            )
        if report.has_may_cycles and strict:
            raise CycleError(
                "May-cycle detected (guarded edges); cannot compute evaluation order in strict mode",
                report.example_may_cycle_path or [],
                is_must_cycle=False,
            )

        exclude: set[NodeKey] = set()
        if report.has_may_cycles and not strict:
            exclude = {n for s in report.may_cycles for n in s}
            warnings.warn(
                f"May-cycles detected; excluding {len(exclude)} nodes from evaluation order",
                UserWarning,
                stacklevel=2,
            )

        self._ensure_csr()
        order: list[NodeKey] = []
        perm: set[NodeKey] = set()
        temp: set[NodeKey] = set()

        def visit(n: NodeKey) -> None:
            if n in perm:
                return
            if n in temp:
                raise CycleError(f"Cycle detected involving {n}", [n], is_must_cycle=True)
            temp.add(n)
            for dep in self._workbook_sorted_keys(self._iter_unguarded_dep_keys(n)):
                if dep in exclude:
                    continue
                if dep in self._nodes and dep not in exclude:
                    visit(dep)
            temp.remove(n)
            perm.add(n)
            order.append(n)

        for key in self._workbook_sorted_keys(self._nodes.keys()):
            if key in exclude:
                continue
            if key not in perm:
                visit(key)

        return order

    def compress_identity_transits(
        self,
        *,
        preserve: set[NodeKey] | None = None,
        record: IdentityTransitCompressionRecord | None = None,
    ) -> list[NodeKey]:
        """Remove identity transit nodes and rewire dependents.

        Transit nodes whose formula is a single cell reference to one dependency
        are removed, dependents' formulas are rewritten, and edges are rewired.
        Requires dependency provenance from graph construction with
        `capture_dependency_provenance=True` for safe edges.

        Identity forwarding rewrites `CellRefNode` leaves only. If a dependent
        mentions the transit as a `RangeNode` endpoint or a whole-column/row
        leaf, the transit is left in place so the formula cannot keep naming a
        removed node and range geometry is not rewritten.

        Skips `is_target` nodes and any keys in `preserve` so public extraction
        and series-bound addresses stay in the graph.

        Node hooks are not invoked for removed or updated nodes. Drops
        `formula_shapes`; callers who want the overlay must rewarm after
        compression. `write_workbook` persists this rewrite into a new
        `.xlsx`; prefer `IdentityTransitCompression().project` when the
        canonical graph must stay unchanged.

        Args:
            preserve: Node keys that must not be forwarded. Always unioned with
                `target_keys()` so marked targets stay public.
            record: When provided, populate with removal lineage for projection
                manifests.

        Returns:
            Keys of removed transit nodes, in removal order.
        """
        from .compression import (
            clear_identity_singleton_ref_cache,
            compression_safe_provenance,
            identity_rewrite_sites_are_cell_refs,
            is_identity_transit,
            require_compression_provenance,
            snapshot_transit_node,
        )

        collapse_preserve = frozenset(self.target_keys())
        if preserve is not None:
            collapse_preserve |= frozenset(normalize_key(key) for key in preserve)

        require_compression_provenance(self)
        self._ensure_staging()
        clear_identity_singleton_ref_cache()
        try:
            heap: list[NodeKey] = list(self._nodes.keys())
            heapq.heapify(heap)
            removed: list[NodeKey] = []
            while heap:
                t_key = heapq.heappop(heap)
                if t_key not in self._nodes:
                    continue
                if t_key in collapse_preserve:
                    continue
                r_key = is_identity_transit(self, t_key)
                if r_key is None:
                    continue
                dependents_t = self._reverse_edges.get(t_key, set())
                if not dependents_t:
                    continue
                ok = True
                for d_key in dependents_t:
                    prov = self._edge_provenance.get((d_key, t_key))
                    if not compression_safe_provenance(prov):
                        ok = False
                        break
                if not ok:
                    continue
                if not identity_rewrite_sites_are_cell_refs(self, t_key, dependents_t):
                    continue

                dependents_before = list(dependents_t)
                snapshot = snapshot_transit_node(self, t_key) if record is not None else None
                self._compress_one_transit(t_key, r_key, record=record)
                if record is not None and snapshot is not None:
                    record.note_removal(t_key, r_key, snapshot)
                removed.append(t_key)
                for d_key in dependents_before:
                    heapq.heappush(heap, d_key)
            self.rebuild_adjacency()
            return removed
        finally:
            clear_identity_singleton_ref_cache()

    def compress_optimal(
        self,
        *,
        preserve: set[NodeKey] | None = None,
        record: OptimalCompressionRecord | None = None,
    ) -> list[NodeKey]:
        """Remove identity transits and inline single-call-site formula nodes.

        Collapses nodes when substitution is safe. Both identity-transit forwarding
        and formula inlining skip `is_target` nodes and any keys in `preserve`
        (external consumers such as series-bound public addresses). Forwarding
        targets are also protected from later inlining. Identity forwarding is
        cell-ref-only: a transit mentioned as a range endpoint or whole-column/row
        leaf is left in place. Drops `formula_shapes`; callers who want the overlay
        must rewarm after compression. `write_workbook` persists this rewrite into
        a new `.xlsx`; prefer `OptimalCompression().project` when the canonical
        graph must stay unchanged.

        Args:
            preserve: Node keys that must not be collapsed (forwarded or inlined).
                Always unioned with `target_keys()` so marked targets stay public.
            record: When provided, populate with removal lineage for projection.

        Returns:
            Keys of removed nodes, in removal order.
        """
        from .compression import (
            IdentityTransitCompressionRecord,
            _incoming_edge_substitutable,
            clear_identity_singleton_ref_cache,
            compression_safe_provenance,
            dependent_context_substitutable,
            identity_rewrite_sites_are_cell_refs,
            is_identity_transit,
            node_body_substitutable,
            require_compression_provenance,
            snapshot_transit_node,
        )

        collapse_preserve = frozenset(self.target_keys())
        if preserve is not None:
            collapse_preserve |= frozenset(normalize_key(key) for key in preserve)
        forwarding_protected: set[NodeKey] = set()

        require_compression_provenance(self)
        self._ensure_staging()
        clear_identity_singleton_ref_cache()
        try:
            heap: list[NodeKey] = list(self._nodes.keys())
            heapq.heapify(heap)
            removed: list[NodeKey] = []
            while heap:
                t_key = heapq.heappop(heap)
                if t_key not in self._nodes:
                    continue
                if t_key in collapse_preserve:
                    continue

                r_key = is_identity_transit(self, t_key)
                if r_key is not None:
                    dependents_t = self._reverse_edges.get(t_key, set())
                    if not dependents_t:
                        continue
                    ok = True
                    for d_key in dependents_t:
                        prov = self._edge_provenance.get((d_key, t_key))
                        if not compression_safe_provenance(prov):
                            ok = False
                            break
                    if not ok:
                        continue
                    if not identity_rewrite_sites_are_cell_refs(self, t_key, dependents_t):
                        continue

                    dependents_before = list(dependents_t)
                    snapshot = snapshot_transit_node(self, t_key) if record is not None else None
                    id_record = IdentityTransitCompressionRecord()
                    self._compress_one_transit(t_key, r_key, record=id_record)
                    if record is not None and snapshot is not None:
                        record.note_forwarding(t_key, r_key, snapshot)
                        record.formula_rewrites.extend(id_record.formula_rewrites)
                    forwarding_protected.add(r_key)
                    removed.append(t_key)
                    for d_key in dependents_before:
                        heapq.heappush(heap, d_key)
                    continue

                t_node = self.get_node(t_key)
                if t_node is None or t_node.is_leaf or not t_node.has_formula:
                    continue
                if t_key in forwarding_protected:
                    continue

                dependents_t = self._reverse_edges.get(t_key, set())
                if len(dependents_t) != 1:
                    continue
                d_key = next(iter(dependents_t))
                if self._is_dependency_reachable(t_key, d_key):
                    continue
                if not _incoming_edge_substitutable(self, d_key, t_key):
                    continue
                if not node_body_substitutable(self, t_key):
                    continue
                if not dependent_context_substitutable(self, d_key, replacing=t_key):
                    continue

                snapshot = snapshot_transit_node(self, t_key) if record is not None else None
                self._inline_one_node(t_key, d_key, record=record)
                if record is not None and snapshot is not None:
                    record.note_inline(t_key, d_key, snapshot)
                removed.append(t_key)
                heapq.heappush(heap, d_key)
            self.rebuild_adjacency()
            return removed
        finally:
            clear_identity_singleton_ref_cache()

    # ---- serialization ------------------------------------------------------

    def __reduce_ex__(self, protocol: SupportsIndex, /) -> str | tuple[Any, ...]:
        """Pickle via a multipart blob so unpickle peak stays near final size."""
        del protocol
        return (loads_graph_blob, (dumps_graph_blob(self),))

    def _invalidate_formula_shapes(self) -> None:
        """Drop interned shapes after a formula rewrite.

        Does not rewarm. Call `warm_formula_shapes` and assign the result if
        the overlay is needed again. A live `FormulaEvaluator` still keeps
        its construction-time shape helpers until reconstructed.
        """
        self.formula_shapes = None

    # ---- internal edge mutation --------------------------------------------

    def _remove_edge(self, from_key: NodeKey, to_key: NodeKey) -> None:
        self._ensure_staging()
        _discard_adjacency(self._edges, from_key, to_key)
        _discard_adjacency(self._reverse_edges, to_key, from_key)
        ek = (from_key, to_key)
        self._guards.pop(ek, None)
        self._edge_provenance.pop(ek, None)

    def _compress_one_transit(
        self,
        t_key: NodeKey,
        r_key: NodeKey,
        *,
        record: IdentityTransitCompressionRecord | None = None,
    ) -> None:
        from .compression import (
            FormulaRewrite,
            direct_provenance_for_key_in_normalized,
            refresh_direct_sites,
            replace_substrings_at_spans,
        )

        for d_key in list(self._reverse_edges.get(t_key, set())):
            prov = self._edge_provenance.get((d_key, t_key))
            guard = self._guards.get((d_key, t_key))
            d_node = self._nodes.get(d_key)
            if d_node is None:
                continue

            before_normalized = d_node.normalized_formula
            new_norm = before_normalized
            if d_node.formula_ast is not None:
                d_node.formula_ast = replace_resolved_cell_ref(
                    d_node.formula_ast,
                    old_key=t_key,
                    new_key=r_key,
                    anchor=d_node.address or d_key,
                )
                new_norm = unparse_normalized_formula(
                    d_node.formula_ast, anchor=d_node.address or d_key
                )
            elif isinstance(prov, EdgeProvenance) and prov.direct_sites_normalized and new_norm:
                new_norm = replace_substrings_at_spans(
                    new_norm, prov.direct_sites_normalized, r_key
                )
            elif new_norm and t_key in new_norm:
                new_norm = new_norm.replace(t_key, r_key)

            if d_node.formula_ast is None and before_normalized != new_norm:
                d_node.apply_formula_text(new_norm)

            if record is not None and before_normalized != new_norm:
                record.formula_rewrites.append(
                    FormulaRewrite(
                        dependent=d_key,
                        before_normalized=before_normalized,
                        after_normalized=new_norm,
                    )
                )

            if before_normalized != new_norm:
                self._invalidate_formula_shapes()

            self._remove_edge(d_key, t_key)
            new_prov = direct_provenance_for_key_in_normalized(new_norm, r_key)
            self.add_edge(d_key, r_key, guard=guard, provenance=new_prov)

            for dep in list(self._edges.get(d_key, set())):
                if dep == r_key:
                    continue
                old_dep_prov = self._edge_provenance.get((d_key, dep))
                if not isinstance(old_dep_prov, EdgeProvenance):
                    continue
                if DependencyCause.direct_ref not in old_dep_prov.causes:
                    continue
                self._edge_provenance[(d_key, dep)] = refresh_direct_sites(
                    old_dep_prov,
                    new_normalized=new_norm,
                    precedent_key=dep,
                )

        for dep in list(self._edges.get(t_key, set())):
            self._remove_edge(t_key, dep)
        self._nodes.pop(t_key, None)
        self._edges.pop(t_key, None)
        self._reverse_edges.pop(t_key, None)

    def _is_dependency_reachable(self, start: NodeKey, target: NodeKey) -> bool:
        """Return whether `target` is reachable from `start` along dependency edges."""
        if start == target:
            return True
        seen: set[NodeKey] = {start}
        stack = list(self._iter_dep_keys(start))
        while stack:
            key = stack.pop()
            if key == target:
                return True
            if key in seen:
                continue
            seen.add(key)
            stack.extend(self._iter_dep_keys(key))
        return False

    def _inline_one_node(
        self,
        t_key: NodeKey,
        d_key: NodeKey,
        *,
        record: OptimalCompressionRecord | None = None,
    ) -> None:
        from .compression import (
            FormulaRewrite,
            direct_provenance_for_key_in_normalized,
            merge_inline_edge_guards,
            refresh_direct_sites,
            substitute_body_at_spans,
        )

        t_node = self._nodes.get(t_key)
        d_node = self._nodes.get(d_key)
        if t_node is None or d_node is None:
            return
        if t_node.formula_ast is None and t_node.normalized_formula is None:
            return

        prov = self._edge_provenance.get((d_key, t_key))
        if not isinstance(prov, EdgeProvenance):
            return

        before_normalized = d_node.normalized_formula
        new_norm = before_normalized
        if d_node.formula_ast is not None and t_node.formula_ast is not None:
            d_node.formula_ast = replace_resolved_cell_ref(
                d_node.formula_ast,
                old_key=t_key,
                new_key=t_key,
                anchor=d_node.address or d_key,
                replacement=bind_axes(t_node.formula_ast, t_node.address or t_key),
            )
            new_norm = unparse_normalized_formula(
                d_node.formula_ast, anchor=d_node.address or d_key
            )
            self._invalidate_formula_shapes()
        elif (
            new_norm is not None
            and t_node.normalized_formula is not None
            and prov.direct_sites_normalized
        ):
            new_norm = substitute_body_at_spans(
                new_norm,
                prov.direct_sites_normalized,
                t_node.normalized_formula,
            )
            if before_normalized != new_norm:
                d_node.apply_formula_text(new_norm)
                self._invalidate_formula_shapes()

        if record is not None and before_normalized != new_norm:
            record.formula_rewrites.append(
                FormulaRewrite(
                    dependent=d_key,
                    before_normalized=before_normalized,
                    after_normalized=new_norm,
                )
            )

        d_other_deps = set(self._edges.get(d_key, set())) - {t_key}
        t_deps = set(self._edges.get(t_key, set()))
        d_dep_guards = {dep: self._guards.get((d_key, dep)) for dep in d_other_deps}
        t_dep_guards = {dep: self._guards.get((t_key, dep)) for dep in t_deps}
        old_dependent_provenance: dict[NodeKey, EdgeProvenance] = {}
        for dep in d_other_deps:
            prov = self.get_edge_attrs(d_key, dep).provenance
            if isinstance(prov, EdgeProvenance):
                old_dependent_provenance[dep] = prov

        for dep in list(self._edges.get(d_key, set())):
            self._remove_edge(d_key, dep)

        inherited_deps = t_deps - d_other_deps
        for dep in d_other_deps | t_deps:
            guard = merge_inline_edge_guards(
                dependent_guard=d_dep_guards.get(dep),
                dependent_has_edge=dep in d_other_deps,
                transit_guard=t_dep_guards.get(dep),
                transit_has_edge=dep in t_deps,
            )
            if dep in inherited_deps:
                new_prov = direct_provenance_for_key_in_normalized(new_norm, dep)
            else:
                old_prov = old_dependent_provenance.get(dep)
                if old_prov is not None:
                    new_prov = refresh_direct_sites(
                        old_prov,
                        new_normalized=new_norm,
                        precedent_key=dep,
                    )
                else:
                    new_prov = direct_provenance_for_key_in_normalized(new_norm, dep)
            self.add_edge(d_key, dep, guard=guard, provenance=new_prov)

        for dep in list(self._edges.get(t_key, set())):
            self._remove_edge(t_key, dep)
        self._nodes.pop(t_key, None)
        self._edges.pop(t_key, None)
        self._reverse_edges.pop(t_key, None)


def _collect_graph_keys(g: DependencyGraph) -> list[str]:
    seen: set[str] = set()

    def add(s: str) -> None:
        seen.add(s)

    for k in g._nodes:
        add(k)
        for d in g._iter_dep_keys(k):
            add(d)
        for d in g._iter_dependent_keys(k):
            add(d)
    for a, b in g._guards:
        add(a)
        add(b)
    for a, b in g._edge_provenance:
        add(a)
        add(b)
    for guard in g._iter_guard_exprs():
        _guard_collect_cellref_keys(guard, add)
    if g.leaf_classification:
        for k in g.leaf_classification:
            add(k)
    return sorted(seen)


def _guard_collect_cellref_keys(expr: GuardExpr, add: Callable[[str], None]) -> None:
    if isinstance(expr, CellRef):
        add(expr.key)
    elif isinstance(expr, Compare):
        _guard_collect_cellref_keys(expr.left, add)
        _guard_collect_cellref_keys(expr.right, add)
    elif isinstance(expr, Not):
        _guard_collect_cellref_keys(expr.operand, add)
    elif isinstance(expr, (And, Or)):
        for o in expr.operands:
            _guard_collect_cellref_keys(o, add)


def _intern_guard_cell_refs(
    expr: GuardExpr,
    keys: list[str],
    *,
    key_index: dict[str, int] | None = None,
) -> GuardExpr:
    rev = key_index if key_index is not None else {s: i for i, s in enumerate(keys)}
    canon = keys

    def ckey(s: str) -> NodeKey:
        return canon[rev[s]]

    def rec(e: GuardExpr) -> GuardExpr:
        if isinstance(e, CellRef):
            return CellRef(key=ckey(e.key))
        if isinstance(e, Compare):
            return Compare(left=rec(e.left), op=e.op, right=rec(e.right))
        if isinstance(e, Not):
            return Not(operand=rec(e.operand))
        if isinstance(e, And):
            return And(operands=tuple(rec(o) for o in e.operands))
        if isinstance(e, Or):
            return Or(operands=tuple(rec(o) for o in e.operands))
        return e

    return intern_guard(rec(expr))


def _scc_cycle_ids(
    n: int,
    neighbors: Callable[[int], Iterable[int]],
) -> list[list[int]]:
    """Return cyclic SCCs as lists of node ids (size>1 or self-loop)."""
    sccs = _tarjan_scc_ids(n, neighbors)
    out: list[list[int]] = []
    for scc in sccs:
        if len(scc) > 1:
            out.append(scc)
            continue
        (v,) = scc
        if any(w == v for w in neighbors(v)):
            out.append(scc)
    return out


def _tarjan_scc_ids(
    n: int,
    neighbors: Callable[[int], Iterable[int]],
) -> list[list[int]]:
    index = 0
    stack: list[int] = []
    on_stack = [False] * n
    indices = [-1] * n
    lowlinks = [0] * n
    result: list[list[int]] = []

    def strongconnect(v: int) -> None:
        nonlocal index
        indices[v] = index
        lowlinks[v] = index
        index += 1
        stack.append(v)
        on_stack[v] = True

        for w in neighbors(v):
            if indices[w] < 0:
                strongconnect(w)
                if lowlinks[w] < lowlinks[v]:
                    lowlinks[v] = lowlinks[w]
            elif on_stack[w] and indices[w] < lowlinks[v]:
                lowlinks[v] = indices[w]

        if lowlinks[v] == indices[v]:
            scc: list[int] = []
            while True:
                w = stack.pop()
                on_stack[w] = False
                scc.append(w)
                if w == v:
                    break
            result.append(scc)

    for v in range(n):
        if indices[v] < 0:
            strongconnect(v)

    return result


def _subgraph_has_cycle_ids(
    nodes: set[int],
    neighbors: Callable[[int], Iterable[int]],
) -> bool:
    if not nodes:
        return False
    ordered = list(nodes)
    local = {node: i for i, node in enumerate(ordered)}

    def local_neighbors(i: int) -> Iterator[int]:
        for w in neighbors(ordered[i]):
            j = local.get(w)
            if j is not None:
                yield j

    return bool(_scc_cycle_ids(len(ordered), local_neighbors))


def _find_cycle_path_ids(
    neighbors: Callable[[int], Iterable[int]],
    keys: Sequence[NodeKey],
    cyclic_sccs: list[list[int]],
) -> list[NodeKey] | None:
    """Find one cycle path within the cyclic SCCs (best-effort)."""
    allowed = {node for scc in cyclic_sccs for node in scc}
    if not allowed:
        return None
    visited: set[int] = set()
    stack: list[int] = []
    in_stack: set[int] = set()

    def dfs(v: int) -> list[int] | None:
        visited.add(v)
        stack.append(v)
        in_stack.add(v)
        for w in neighbors(v):
            if w not in allowed:
                continue
            if w in in_stack:
                i = stack.index(w)
                return stack[i:] + [w]
            if w not in visited:
                out = dfs(w)
                if out is not None:
                    return out
        stack.pop()
        in_stack.remove(v)
        return None

    for start in allowed:
        if start not in visited:
            path = dfs(start)
            if path is not None:
                return [keys[i] for i in path]
    return None


def _apply_guard_constraints(
    constraints: GuardConstraints,
    guard: GuardExpr | None,
    *,
    cell_type_env: CellTypeEnv | None = None,
    aliases: Mapping[NodeKey, NodeKey] | None = None,
) -> list[GuardConstraints]:
    """Conjoin an edge guard onto the current constraints.

    For disjunctive guards (OR), this returns multiple possible constraint sets,
    one per feasible disjunct (best-effort). This keeps cycle feasibility checks
    conservative without requiring full boolean reasoning.
    """
    if guard is not None and aliases:
        guard = rewrite_guard_aliases(guard, aliases)
    if guard is None:
        return [constraints]
    if isinstance(guard, Or):
        out: list[GuardConstraints] = []
        # Best-effort: branch on each disjunct and keep feasible ones.
        for g in guard.operands:
            nxt = constraints.add(g, cell_type_env=cell_type_env)
            if nxt is None:
                continue
            out.append(nxt)
            # Avoid pathological blow-ups.
            if len(out) >= 32:
                break
        return out
    nxt = constraints.add(guard, cell_type_env=cell_type_env)
    return [] if nxt is None else [nxt]


def _seed_guard_constraints(cell_type_env: CellTypeEnv | None) -> GuardConstraints:
    seed = GuardConstraints()
    if cell_type_env is None:
        return seed
    seeded = seed.seed_cell_type_env(cell_type_env)
    return seed if seeded is None else seeded


def _subgraph_has_feasible_cycle(
    graph: DependencyGraph,
    nodes: set[NodeKey],
    *,
    cell_type_env: CellTypeEnv | None = None,
    aliases: Mapping[NodeKey, NodeKey] | None = None,
) -> bool:
    """Return whether `nodes` contains a guard-feasible cycle.

    True when at least one cycle within `nodes` has jointly consistent
    accumulated edge guards (symbolic, no evaluation).
    """
    visited: set[tuple[NodeKey, GuardConstraints]] = set()
    on_stack: set[NodeKey] = set()

    def dfs(v: NodeKey, c: GuardConstraints) -> bool:
        state = (v, c)
        if state in visited:
            return False
        visited.add(state)
        on_stack.add(v)

        for raw_w in graph._iter_dep_keys(v):
            w = graph._resolve_graph_endpoint(raw_w)
            if w is None or w not in nodes:
                continue
            guard = graph._stored_guard(v, raw_w)
            if guard is None:
                guard = graph._stored_guard(v, w)
            for c2 in _apply_guard_constraints(
                c, guard, cell_type_env=cell_type_env, aliases=aliases
            ):
                if w in on_stack:
                    return True
                if dfs(w, c2):
                    return True

        on_stack.remove(v)
        return False

    seed = _seed_guard_constraints(cell_type_env)
    return any(dfs(n, seed) for n in nodes)


def _find_feasible_cycle_path(
    graph: DependencyGraph,
    nodes: set[NodeKey],
    *,
    cell_type_env: CellTypeEnv | None = None,
    aliases: Mapping[NodeKey, NodeKey] | None = None,
) -> list[NodeKey] | None:
    """Best-effort: find one feasible cycle path within `nodes` (symbolic constraints)."""
    visited: set[tuple[NodeKey, GuardConstraints]] = set()
    stack: list[NodeKey] = []
    on_stack: set[NodeKey] = set()

    def dfs(v: NodeKey, c: GuardConstraints) -> list[NodeKey] | None:
        state = (v, c)
        if state in visited:
            return None
        visited.add(state)
        stack.append(v)
        on_stack.add(v)

        for raw_w in graph._iter_dep_keys(v):
            w = graph._resolve_graph_endpoint(raw_w)
            if w is None or w not in nodes:
                continue
            guard = graph._stored_guard(v, raw_w)
            if guard is None:
                guard = graph._stored_guard(v, w)
            for c2 in _apply_guard_constraints(
                c, guard, cell_type_env=cell_type_env, aliases=aliases
            ):
                if w in on_stack:
                    i = stack.index(w)
                    return stack[i:] + [w]
                out = dfs(w, c2)
                if out is not None:
                    return out

        stack.pop()
        on_stack.remove(v)
        return None

    seed = _seed_guard_constraints(cell_type_env)
    for n in nodes:
        out = dfs(n, seed)
        if out is not None:
            return out
    return None
