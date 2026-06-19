from __future__ import annotations

import copy
import heapq
import warnings
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal, Protocol, runtime_checkable

if TYPE_CHECKING:
    from .compression import IdentityTransitCompressionRecord, OptimalCompressionRecord

from excel_grapher.core.address_keys import normalize_key, sort_node_keys

from .dependency_provenance import DependencyCause, EdgeProvenance, merge_edge_provenance
from .guard import And, CellRef, Compare, GuardConstraints, GuardExpr, Not, Or, or_guard
from .node import Node, NodeKey, NodeView, node_to_view

NodeHook = Callable[[NodeKey, Node], None]

EdgeKey = tuple[NodeKey, NodeKey]

_PICKLE_VERSION = 2


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

    Consumers that only read a graph (for example `to_networkx` and
    `CodeGenerator`) can accept any object satisfying this protocol, including
    projected facades such as `ProjectionResult`, without depending on the
    concrete `DependencyGraph` type. It captures node iteration, node and edge
    lookups, key listings, leaf/formula/target classification, and evaluation
    order; mutation is intentionally excluded.
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

    def get_edge_attrs(self, from_key: NodeKey, to_key: NodeKey) -> EdgeAttrs: ...

    def get_edge_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None: ...

    def leaf_keys(self) -> list[NodeKey]: ...

    def formula_keys(self) -> list[NodeKey]: ...

    def target_keys(self) -> list[NodeKey]: ...

    def evaluation_order(
        self, *, strict: bool = ..., iterate_enabled: bool | None = ...
    ) -> list[NodeKey]: ...

    def cycle_report(self) -> CycleReport: ...


@dataclass
class DependencyGraph:
    _nodes: dict[NodeKey, Node] = field(default_factory=dict)
    _edges: dict[NodeKey, set[NodeKey]] = field(default_factory=dict)  # node -> deps
    _reverse_edges: dict[NodeKey, set[NodeKey]] = field(default_factory=dict)  # node -> dependents
    _guards: dict[EdgeKey, GuardExpr] = field(default_factory=dict)
    _edge_extra: dict[EdgeKey, dict[str, Any]] = field(default_factory=dict)
    _hooks: list[NodeHook] = field(default_factory=list)
    leaf_classification: dict[str, str] | None = None
    sheet_order: list[str] | None = None
    sheet_bounds: dict[str, tuple[int, int]] | None = None
    named_ranges: dict[str, tuple[str, str]] | None = None
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None

    def copy(self) -> DependencyGraph:
        """Return a deep copy of this graph (node hooks are not copied)."""
        cloned = copy.deepcopy(self)
        cloned._hooks = []
        return cloned

    # ---- node insertion and iteration ---------------------------------------

    def add_node(self, node: Node) -> None:
        key = node.key
        self._nodes[key] = node
        self._edges.setdefault(key, set())
        self._reverse_edges.setdefault(key, set())
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
        **attrs: Any,
    ) -> None:
        """Add edge: from_key depends on to_key (from_key -> to_key)."""
        from_key = normalize_key(from_key)
        to_key = normalize_key(to_key)
        unknown_attrs = [k for k in attrs if k != "provenance"]
        if unknown_attrs:
            names = ", ".join(sorted(unknown_attrs))
            raise ValueError(f"Unsupported edge attrs: {names}")
        ek = (from_key, to_key)
        deps_existing = self._edges.get(from_key)
        was_present = deps_existing is not None and to_key in deps_existing

        self._edges.setdefault(from_key, set()).add(to_key)
        self._reverse_edges.setdefault(to_key, set()).add(from_key)

        if not was_present:
            merged_guard = guard
            merged_extra: dict[str, Any] = {}
        else:
            existing_guard = self._guards.get(ek)
            if existing_guard is None or guard is None:
                merged_guard = None
            elif existing_guard == guard:
                merged_guard = guard
            else:
                merged_guard = or_guard(existing_guard, guard)
            merged_extra = dict(self._edge_extra.get(ek, {}))

        merged_extra.update({k: v for k, v in attrs.items() if k != "provenance"})
        prov_new = attrs.get("provenance")
        if prov_new is not None and isinstance(prov_new, EdgeProvenance):
            old_prov = merged_extra.get("provenance")
            merged_extra["provenance"] = merge_edge_provenance(
                old_prov if isinstance(old_prov, EdgeProvenance) else None,
                prov_new,
            )

        if merged_guard is not None:
            self._guards[ek] = merged_guard
        else:
            self._guards.pop(ek, None)

        if merged_extra:
            self._edge_extra[ek] = merged_extra
        else:
            self._edge_extra.pop(ek, None)

    # ---- public read API ----------------------------------------------------

    def get_node(self, key: NodeKey) -> NodeView | None:
        """Return an immutable `NodeView` snapshot, or `None` if missing."""
        node = self._nodes.get(normalize_key(key))
        if node is None:
            return None
        return node_to_view(node)

    def get_dependencies(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return an immutable snapshot of `key`'s dependencies (cells it reads)."""
        deps = self._edges.get(normalize_key(key))
        if not deps:
            return frozenset()
        return frozenset(deps)

    def get_dependents(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return an immutable snapshot of cells that depend on `key`."""
        deps = self._reverse_edges.get(normalize_key(key))
        if not deps:
            return frozenset()
        return frozenset(deps)

    def get_edge_attrs(self, from_key: NodeKey, to_key: NodeKey) -> EdgeAttrs:
        """Return a typed snapshot of the attributes on edge `from_key -> to_key`.

        When the edge does not exist, returns an `EdgeAttrs` with all fields
        set to `None`.
        """
        fk = normalize_key(from_key)
        tk = normalize_key(to_key)
        if tk not in self._edges.get(fk, set()):
            return EdgeAttrs()
        extra = self._edge_extra.get((fk, tk), {})
        prov = extra.get("provenance")
        return EdgeAttrs(
            guard=self._guards.get((fk, tk)),
            provenance=prov if isinstance(prov, EdgeProvenance) else None,
        )

    def get_edge_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None:
        """Return the guard on edge `from_key -> to_key`, or `None` if none."""
        fk = normalize_key(from_key)
        tk = normalize_key(to_key)
        v = self._guards.get((fk, tk))
        return v if isinstance(v, GuardExpr) else None

    # ---- durable node mutation ---------------------------------------------

    def set_node_value(self, key: NodeKey, value: Any) -> None:
        """Set a node's `value` field durably. Raises `KeyError` if missing."""
        nk = normalize_key(key)
        node = self._nodes.get(nk)
        if node is None:
            raise KeyError(f"Cell {key} not found in graph")
        node.value = value

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
        node.metadata = dict(metadata)

    def set_node_formula(
        self,
        key: NodeKey,
        formula: str | None,
        normalized_formula: str | None,
    ) -> None:
        """Set a node's `formula` and `normalized_formula` durably.

        Edges are not recomputed; callers rewiring dependencies must update edges
        explicitly. Intended for projection authors building export-only graph
        views. Raises `KeyError` if the node is missing.
        """
        nk = normalize_key(key)
        node = self._nodes.get(nk)
        if node is None:
            raise KeyError(f"Cell {key} not found in graph")
        node.formula = formula
        node.normalized_formula = normalized_formula

    def remove_node(self, key: NodeKey) -> None:
        """Remove a node and all of its incident edges.

        Both outgoing dependency edges and incoming dependent edges are dropped,
        along with their guards and provenance. Dependent formulas are not
        rewritten; callers collapsing nodes must update dependents explicitly.
        Node hooks are not invoked. No-op if the node is absent.
        """
        nk = normalize_key(key)
        for dep in list(self._edges.get(nk, set())):
            self._remove_edge(nk, dep)
        for dependent in list(self._reverse_edges.get(nk, set())):
            self._remove_edge(dependent, nk)
        self._nodes.pop(nk, None)
        self._edges.pop(nk, None)
        self._reverse_edges.pop(nk, None)

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
        """Iterate over (key, node) pairs for nodes that contain formulas."""
        for key, node in self._nodes.items():
            if node.formula is not None:
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
            source=(k for k, node in self._nodes.items() if node.formula is not None),
        )

    def leaf_keys(self) -> list[NodeKey]:
        """Return sorted list of keys for nodes with no dependency edges (leaves)."""
        return self.keys(
            order="workbook", source=(k for k, node in self._nodes.items() if node.is_leaf)
        )

    def target_keys(self) -> list[NodeKey]:
        """Return sorted list of keys marked as original build targets."""
        return self.keys(
            order="workbook", source=(k for k, node in self._nodes.items() if node.is_target)
        )

    def roots(self) -> Iterator[NodeKey]:
        for key in self._nodes:
            if not self._reverse_edges.get(key):
                yield key

    # ---- adjacency helpers for cycle/order analysis ------------------------

    def _unconditional_adjacency(self) -> dict[NodeKey, set[NodeKey]]:
        out: dict[NodeKey, set[NodeKey]] = {k: set() for k in self._nodes}
        for k in self._nodes:
            for dep in self._edges.get(k, ()):
                if dep not in self._nodes:
                    continue
                if (k, dep) not in self._guards:
                    out[k].add(dep)
        return out

    def _all_adjacency(self) -> dict[NodeKey, set[NodeKey]]:
        out: dict[NodeKey, set[NodeKey]] = {k: set() for k in self._nodes}
        for k in self._nodes:
            for dep in self._edges.get(k, ()):
                if dep not in self._nodes:
                    continue
                out[k].add(dep)
        return out

    def cycle_report(self) -> CycleReport:
        uncond = self._unconditional_adjacency()
        all_edges = self._all_adjacency()

        must_sccs = _scc_cycles(uncond)
        must_nodes = {n for s in must_sccs for n in s}
        example_must = _find_cycle_path(uncond, must_nodes) if must_sccs else None

        may_sccs: list[set[NodeKey]] = []
        example_may: list[NodeKey] | None = None
        for scc in _scc_cycles(all_edges):
            # If this SCC already has an unconditional cycle, it's not "may".
            if _subgraph_has_cycle(uncond, scc):
                continue
            # Filter out SCCs whose only cycles are infeasible due to contradictory guards.
            if not _subgraph_has_feasible_cycle(self, scc):
                continue
            may_sccs.append(scc)

        if may_sccs:
            # Best-effort: find a feasible example path inside the first may-SCC.
            example_may = _find_feasible_cycle_path(self, may_sccs[0])

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
        self, *, strict: bool = True, iterate_enabled: bool | None = None
    ) -> list[NodeKey]:
        """Return nodes in dependency-first order (leaves before formulas that use them).

        Edge direction is A -> B meaning A depends on B. This method returns an
        ordering suitable for sequential evaluation (dependencies first).

        If `iterate_enabled` is True (workbook has iterative calculation on), any
        must-cycle or may-cycle is rejected: generated Python does not emulate Excel's
        iterative convergence. Pass `False` or `None` to apply the usual strict /
        non-strict rules without this check.
        """
        report = self.cycle_report()
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

        adjacency = self._unconditional_adjacency()
        order: list[NodeKey] = []
        perm: set[NodeKey] = set()
        temp: set[NodeKey] = set()

        def visit(n: NodeKey) -> None:
            if n in perm:
                return
            if n in temp:
                raise CycleError(f"Cycle detected involving {n}", [n], is_must_cycle=True)
            temp.add(n)
            for dep in self._workbook_sorted_keys(adjacency.get(n, set())):
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
        record: IdentityTransitCompressionRecord | None = None,
    ) -> list[NodeKey]:
        """Remove identity transit nodes and rewire dependents.

        Transit nodes whose formula is a single cell reference to one dependency
        are removed, dependents' formulas are rewritten, and edges are rewired.
        Requires dependency provenance from graph construction with
        `capture_dependency_provenance=True` for safe edges.

        Node hooks are not invoked for removed or updated nodes.

        Args:
            record: When provided, populate with removal lineage for projection
                manifests.

        Returns:
            Keys of removed transit nodes, in removal order.
        """
        from .compression import (
            clear_identity_singleton_ref_cache,
            compression_safe_provenance,
            is_identity_transit,
            require_compression_provenance,
            snapshot_transit_node,
        )

        require_compression_provenance(self)
        clear_identity_singleton_ref_cache()
        try:
            heap: list[NodeKey] = list(self._nodes.keys())
            heapq.heapify(heap)
            removed: list[NodeKey] = []
            while heap:
                t_key = heapq.heappop(heap)
                if t_key not in self._nodes:
                    continue
                r_key = is_identity_transit(self, t_key)
                if r_key is None:
                    continue
                dependents_t = self._reverse_edges.get(t_key, set())
                if not dependents_t:
                    continue
                ok = True
                for d_key in dependents_t:
                    prov = self._edge_extra.get((d_key, t_key), {}).get("provenance")
                    if not compression_safe_provenance(
                        prov if isinstance(prov, EdgeProvenance) else None
                    ):
                        ok = False
                        break
                if not ok:
                    continue

                dependents_before = list(dependents_t)
                snapshot = snapshot_transit_node(self, t_key) if record is not None else None
                self._compress_one_transit(t_key, r_key, record=record)
                if record is not None and snapshot is not None:
                    record.note_removal(t_key, r_key, snapshot)
                removed.append(t_key)
                for d_key in dependents_before:
                    heapq.heappush(heap, d_key)
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

        Collapses nodes with exactly one dependent when substitution is safe.
        Identity transit forwarding is always attempted; formula inlining respects
        `preserve` (defaults to target-marked nodes) and protects forwarding targets
        from later inlining.

        Args:
            preserve: Node keys that must not be inlined into their dependent.
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
            is_identity_transit,
            node_body_substitutable,
            require_compression_provenance,
            snapshot_transit_node,
        )

        if preserve is None:
            inline_preserve = frozenset(self.target_keys())
        else:
            inline_preserve = frozenset(normalize_key(key) for key in preserve)
        forwarding_protected: set[NodeKey] = set()

        require_compression_provenance(self)
        clear_identity_singleton_ref_cache()
        try:
            heap: list[NodeKey] = list(self._nodes.keys())
            heapq.heapify(heap)
            removed: list[NodeKey] = []
            while heap:
                t_key = heapq.heappop(heap)
                if t_key not in self._nodes:
                    continue

                r_key = is_identity_transit(self, t_key)
                if r_key is not None:
                    dependents_t = self._reverse_edges.get(t_key, set())
                    if not dependents_t:
                        continue
                    ok = True
                    for d_key in dependents_t:
                        prov = self._edge_extra.get((d_key, t_key), {}).get("provenance")
                        if not compression_safe_provenance(
                            prov if isinstance(prov, EdgeProvenance) else None
                        ):
                            ok = False
                            break
                    if not ok:
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
                if t_node is None or t_node.is_leaf or t_node.formula is None:
                    continue
                if t_key in inline_preserve or t_key in forwarding_protected:
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
            return removed
        finally:
            clear_identity_singleton_ref_cache()

    # ---- serialization ------------------------------------------------------

    def __getstate__(self) -> dict[str, Any]:
        keys_sorted = _collect_graph_keys(self)
        idx = {k: i for i, k in enumerate(keys_sorted)}
        return {
            "v": _PICKLE_VERSION,
            "keys": keys_sorted,
            "_nodes": {idx[k]: n for k, n in self._nodes.items()},
            "_edges": {idx[k]: {idx[d] for d in ds} for k, ds in self._edges.items()},
            "_reverse_edges": {
                idx[k]: {idx[d] for d in ds} for k, ds in self._reverse_edges.items()
            },
            "_guards": [(idx[a], idx[b], g) for (a, b), g in self._guards.items()],
            "_edge_extra": [(idx[a], idx[b], dict(e)) for (a, b), e in self._edge_extra.items()],
            "_hooks": self._hooks,
            "leaf_classification": self.leaf_classification,
            "sheet_order": list(self.sheet_order) if self.sheet_order is not None else None,
            "named_ranges": dict(self.named_ranges) if self.named_ranges else None,
            "named_range_ranges": (
                dict(self.named_range_ranges) if self.named_range_ranges else None
            ),
        }

    def __setstate__(self, state: dict[str, Any]) -> None:
        if not isinstance(state, dict) or state.get("v") != _PICKLE_VERSION:
            raise TypeError(
                "Unsupported or corrupted DependencyGraph pickle; rebuild the graph cache."
            )
        keys = state["keys"]
        key_index = {s: i for i, s in enumerate(keys)}

        def i2k(i: int) -> NodeKey:
            return keys[i]

        self._nodes = {i2k(i): n for i, n in state["_nodes"].items()}
        self._edges = {i2k(i): {i2k(d) for d in ds} for i, ds in state["_edges"].items()}
        self._reverse_edges = {
            i2k(i): {i2k(d) for d in ds} for i, ds in state["_reverse_edges"].items()
        }
        self._guards = {
            (i2k(a), i2k(b)): _intern_guard_cell_refs(g, keys, key_index=key_index)
            for a, b, g in state["_guards"]
        }
        self._edge_extra = {(i2k(a), i2k(b)): dict(e) for a, b, e in state["_edge_extra"]}
        self._hooks = state["_hooks"]
        lc = state["leaf_classification"]
        if lc:
            self.leaf_classification = {keys[key_index[k]]: v for k, v in lc.items()}
        else:
            self.leaf_classification = None
        sheet_order = state.get("sheet_order")
        if sheet_order:
            self.sheet_order = list(sheet_order)
        else:
            self.sheet_order = None
        nr = state.get("named_ranges")
        self.named_ranges = dict(nr) if nr else None
        nrr = state.get("named_range_ranges")
        self.named_range_ranges = dict(nrr) if nrr else None

    # ---- internal edge mutation --------------------------------------------

    def _remove_edge(self, from_key: NodeKey, to_key: NodeKey) -> None:
        self._edges.setdefault(from_key, set()).discard(to_key)
        self._reverse_edges.setdefault(to_key, set()).discard(from_key)
        ek = (from_key, to_key)
        self._guards.pop(ek, None)
        self._edge_extra.pop(ek, None)

    def _compress_one_transit(
        self,
        t_key: NodeKey,
        r_key: NodeKey,
        *,
        record: IdentityTransitCompressionRecord | None = None,
    ) -> None:
        from .compression import (
            FormulaRewrite,
            direct_provenance_for_key_in_strings,
            refresh_direct_sites,
            replace_substrings_at_spans,
        )

        for d_key in list(self._reverse_edges.get(t_key, set())):
            extra = self._edge_extra.get((d_key, t_key), {})
            prov = extra.get("provenance")
            guard = self._guards.get((d_key, t_key))
            d_node = self._nodes.get(d_key)
            if d_node is None:
                continue

            before_formula = d_node.formula
            before_normalized = d_node.normalized_formula
            new_formula = before_formula
            new_norm = before_normalized
            if isinstance(prov, EdgeProvenance) and prov.direct_sites_formula and new_formula:
                new_formula = replace_substrings_at_spans(
                    new_formula, prov.direct_sites_formula, r_key
                )
            elif new_formula and t_key in new_formula:
                new_formula = new_formula.replace(t_key, r_key)

            if isinstance(prov, EdgeProvenance) and prov.direct_sites_normalized and new_norm:
                new_norm = replace_substrings_at_spans(
                    new_norm, prov.direct_sites_normalized, r_key
                )
            elif new_norm and t_key in new_norm:
                new_norm = new_norm.replace(t_key, r_key)

            if record is not None and (
                before_formula != new_formula or before_normalized != new_norm
            ):
                record.formula_rewrites.append(
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

            self._remove_edge(d_key, t_key)
            new_prov = direct_provenance_for_key_in_strings(new_formula, new_norm, r_key)
            self.add_edge(d_key, r_key, guard=guard, provenance=new_prov)

            for dep in list(self._edges.get(d_key, set())):
                if dep == r_key:
                    continue
                dep_extra = self._edge_extra.get((d_key, dep), {})
                old_dep_prov = dep_extra.get("provenance")
                if not isinstance(old_dep_prov, EdgeProvenance):
                    continue
                if DependencyCause.direct_ref not in old_dep_prov.causes:
                    continue
                dep_extra["provenance"] = refresh_direct_sites(
                    old_dep_prov,
                    old_formula=before_formula,
                    new_formula=new_formula,
                    old_normalized=before_normalized,
                    new_normalized=new_norm,
                    precedent_key=dep,
                )
                self._edge_extra[(d_key, dep)] = dep_extra

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
        stack = list(self._edges.get(start, set()))
        while stack:
            key = stack.pop()
            if key == target:
                return True
            if key in seen:
                continue
            seen.add(key)
            stack.extend(self._edges.get(key, set()))
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
            direct_provenance_for_key_in_strings,
            refresh_direct_sites,
            substitute_body_at_spans,
        )

        t_node = self._nodes.get(t_key)
        d_node = self._nodes.get(d_key)
        if t_node is None or d_node is None:
            return
        if t_node.formula is None or t_node.normalized_formula is None:
            return

        extra = self._edge_extra.get((d_key, t_key), {})
        prov = extra.get("provenance")
        if not isinstance(prov, EdgeProvenance):
            return

        before_formula = d_node.formula
        before_normalized = d_node.normalized_formula
        new_formula = before_formula
        new_norm = before_normalized
        if new_formula is not None and prov.direct_sites_formula:
            new_formula = substitute_body_at_spans(
                new_formula,
                prov.direct_sites_formula,
                t_node.formula,
            )
        if new_norm is not None and prov.direct_sites_normalized:
            new_norm = substitute_body_at_spans(
                new_norm,
                prov.direct_sites_normalized,
                t_node.normalized_formula,
            )

        if record is not None and (before_formula != new_formula or before_normalized != new_norm):
            record.formula_rewrites.append(
                FormulaRewrite(
                    dependent=d_key,
                    before_formula=before_formula,
                    after_formula=new_formula,
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

        d_node.formula = new_formula
        d_node.normalized_formula = new_norm

        for dep in list(self._edges.get(d_key, set())):
            self._remove_edge(d_key, dep)

        inherited_deps = t_deps - d_other_deps
        for dep in d_other_deps | t_deps:
            guard = d_dep_guards.get(dep)
            if guard is None:
                guard = t_dep_guards.get(dep)
            if dep in inherited_deps:
                new_prov = direct_provenance_for_key_in_strings(new_formula, new_norm, dep)
            else:
                old_prov = old_dependent_provenance.get(dep)
                if old_prov is not None:
                    new_prov = refresh_direct_sites(
                        old_prov,
                        old_formula=before_formula,
                        new_formula=new_formula,
                        old_normalized=before_normalized,
                        new_normalized=new_norm,
                        precedent_key=dep,
                    )
                else:
                    new_prov = direct_provenance_for_key_in_strings(new_formula, new_norm, dep)
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
    for k, deps in g._edges.items():
        add(k)
        for d in deps:
            add(d)
    for k, deps in g._reverse_edges.items():
        add(k)
        for d in deps:
            add(d)
    for a, b in g._guards:
        add(a)
        add(b)
    for a, b in g._edge_extra:
        add(a)
        add(b)
    for guard in g._guards.values():
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

    return rec(expr)


def _scc_cycles(adj: dict[NodeKey, set[NodeKey]]) -> list[set[NodeKey]]:
    """Return SCCs that are cyclic (size>1 or self-loop)."""
    sccs = _tarjan_scc(adj)
    out: list[set[NodeKey]] = []
    for scc in sccs:
        if len(scc) > 1:
            out.append(scc)
        else:
            (n,) = tuple(scc)
            if n in adj.get(n, set()):
                out.append(scc)
    return out


def _tarjan_scc(adj: dict[NodeKey, set[NodeKey]]) -> list[set[NodeKey]]:
    index = 0
    stack: list[NodeKey] = []
    on_stack: set[NodeKey] = set()
    indices: dict[NodeKey, int] = {}
    lowlinks: dict[NodeKey, int] = {}
    result: list[set[NodeKey]] = []

    def strongconnect(v: NodeKey) -> None:
        nonlocal index
        indices[v] = index
        lowlinks[v] = index
        index += 1
        stack.append(v)
        on_stack.add(v)

        for w in adj.get(v, set()):
            if w not in indices:
                strongconnect(w)
                lowlinks[v] = min(lowlinks[v], lowlinks[w])
            elif w in on_stack:
                lowlinks[v] = min(lowlinks[v], indices[w])

        if lowlinks[v] == indices[v]:
            scc: set[NodeKey] = set()
            while True:
                w = stack.pop()
                on_stack.remove(w)
                scc.add(w)
                if w == v:
                    break
            result.append(scc)

    for v in adj:
        if v not in indices:
            strongconnect(v)

    return result


def _subgraph_has_cycle(adj: dict[NodeKey, set[NodeKey]], nodes: set[NodeKey]) -> bool:
    sub = {n: {d for d in adj.get(n, set()) if d in nodes} for n in nodes}
    return bool(_scc_cycles(sub))


def _find_cycle_path(adj: dict[NodeKey, set[NodeKey]], nodes: set[NodeKey]) -> list[NodeKey] | None:
    """Find one cycle path within the given node subset (best-effort)."""
    visited: set[NodeKey] = set()
    stack: list[NodeKey] = []
    in_stack: set[NodeKey] = set()

    def dfs(v: NodeKey) -> list[NodeKey] | None:
        visited.add(v)
        stack.append(v)
        in_stack.add(v)
        for w in adj.get(v, set()):
            if w not in nodes:
                continue
            if w in in_stack:
                # Return the cycle portion from w to v (inclusive) plus w to close.
                i = stack.index(w)
                return stack[i:] + [w]
            if w not in visited:
                out = dfs(w)
                if out is not None:
                    return out
        stack.pop()
        in_stack.remove(v)
        return None

    for n in nodes:
        if n not in visited:
            p = dfs(n)
            if p is not None:
                return p
    return None


def _apply_guard_constraints(
    constraints: GuardConstraints, guard: GuardExpr | None
) -> list[GuardConstraints]:
    """Conjoin an edge guard onto the current constraints.

    For disjunctive guards (OR), this returns multiple possible constraint sets,
    one per feasible disjunct (best-effort). This keeps cycle feasibility checks
    conservative without requiring full boolean reasoning.
    """
    if guard is None:
        return [constraints]
    if isinstance(guard, Or):
        out: list[GuardConstraints] = []
        # Best-effort: branch on each disjunct and keep feasible ones.
        for g in guard.operands:
            nxt = constraints.add(g)
            if nxt is None:
                continue
            out.append(nxt)
            # Avoid pathological blow-ups.
            if len(out) >= 32:
                break
        return out
    nxt = constraints.add(guard)
    return [] if nxt is None else [nxt]


def _subgraph_has_feasible_cycle(graph: DependencyGraph, nodes: set[NodeKey]) -> bool:
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

        for w in graph._edges.get(v, ()):
            if w not in nodes:
                continue
            guard = graph._guards.get((v, w))
            for c2 in _apply_guard_constraints(c, guard):
                if w in on_stack:
                    return True
                if dfs(w, c2):
                    return True

        on_stack.remove(v)
        return False

    seed = GuardConstraints()
    return any(dfs(n, seed) for n in nodes)


def _find_feasible_cycle_path(graph: DependencyGraph, nodes: set[NodeKey]) -> list[NodeKey] | None:
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

        for w in graph._edges.get(v, ()):
            if w not in nodes:
                continue
            guard = graph._guards.get((v, w))
            for c2 in _apply_guard_constraints(c, guard):
                if w in on_stack:
                    i = stack.index(w)
                    return stack[i:] + [w]
                out = dfs(w, c2)
                if out is not None:
                    return out

        stack.pop()
        on_stack.remove(v)
        return None

    seed = GuardConstraints()
    for n in nodes:
        out = dfs(n, seed)
        if out is not None:
            return out
    return None
