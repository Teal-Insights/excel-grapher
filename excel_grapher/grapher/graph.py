from __future__ import annotations

import copy
import heapq
import warnings
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal, Protocol, SupportsIndex, runtime_checkable

if TYPE_CHECKING:
    from .compression import IdentityTransitCompressionRecord, OptimalCompressionRecord

from excel_grapher.core.address_keys import (
    CellKey,
    normalize_key,
    parse_node_key,
    sort_node_keys,
)
from excel_grapher.core.formula_ast import AstNode

from .dependency_provenance import DependencyCause, EdgeProvenance, merge_edge_provenance
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
)
from .node import Node, NodeKey, NodeView, copy_metadata, member_keys, node_to_view

NodeHook = Callable[[NodeKey, Node], None]

EdgeKey = tuple[NodeKey, NodeKey]

_PICKLE_VERSION = 3


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

    def resolve_endpoint(self, address: NodeKey) -> NodeKey | None: ...

    def get_dependency_nodes(self, address: NodeKey) -> frozenset[NodeKey]: ...

    def get_edge_attrs(self, from_key: NodeKey, to_key: NodeKey) -> EdgeAttrs: ...

    def get_edge_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None: ...

    def is_guarded(self, from_key: NodeKey, to_key: NodeKey) -> bool: ...

    def leaf_keys(self) -> list[NodeKey]: ...

    def formula_keys(self) -> list[NodeKey]: ...

    def target_keys(self) -> list[NodeKey]: ...

    def evaluation_order(
        self, *, strict: bool = ..., iterate_enabled: bool | None = ...
    ) -> list[NodeKey]: ...

    def cycle_report(self) -> CycleReport: ...


@dataclass
class DependencyGraph:
    """Mutable workbook dependency graph.

    Node identity is the canonical address string (`CellKey` / `RangeKey` /
    `UnionKey`). `get_node(key)` is exact-key only — a member cell of a
    multi-cell node is not stored under its own key. Prefer `locate_cell` (or
    `cell_owner`) when resolving a workbook cell that may belong to a group;
    then call `get_node` on the returned `node_key`.

    Edge endpoints may name member cells. `get_dependencies` /
    `get_dependents` / `get_edge_attrs` keep those raw keys so member-level
    provenance is preserved. Use `resolve_endpoint` or `get_dependency_nodes`
    when a stored graph node is required (evaluation order, export, codegen).
    """

    _nodes: dict[NodeKey, Node] = field(default_factory=dict)
    _edges: dict[NodeKey, set[NodeKey]] = field(default_factory=dict)  # node -> deps
    _reverse_edges: dict[NodeKey, set[NodeKey]] = field(default_factory=dict)  # node -> dependents
    _guards: dict[EdgeKey, GuardExpr] = field(default_factory=dict)
    _edge_provenance: dict[EdgeKey, EdgeProvenance] = field(default_factory=dict)
    _hooks: list[NodeHook] = field(default_factory=list)
    # Cell key -> owning node key (cell owns itself; multi-cell owns members).
    _occupancy: dict[NodeKey, NodeKey] = field(default_factory=dict)
    leaf_classification: dict[str, str] | None = None
    sheet_order: list[str] | None = None
    sheet_bounds: dict[str, tuple[int, int]] | None = None
    named_ranges: dict[str, tuple[str, str]] | None = None
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None
    # Opt-in AST cache from warm_ast_cache; not JSON-serialized; re-warm after load.
    preparsed_formulas: dict[str, AstNode] | None = None

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
        cloned._nodes = {
            key: Node(
                sheet=node.sheet,
                column=node.column,
                row=node.row,
                formula=node.formula,
                normalized_formula=node.normalized_formula,
                value=node.value,
                is_leaf=node.is_leaf,
                is_target=node.is_target,
                metadata=copy_metadata(node.metadata),
                kind=node.kind,
                min_col=node.min_col,
                min_row=node.min_row,
                max_col=node.max_col,
                max_row=node.max_row,
                address=node.address,
            )
            for key, node in self._nodes.items()
        }
        cloned._edges = {key: set(deps) for key, deps in self._edges.items()}
        cloned._reverse_edges = {
            key: set(dependents) for key, dependents in self._reverse_edges.items()
        }
        cloned._guards = dict(self._guards)
        cloned._edge_provenance = dict(self._edge_provenance)
        cloned._occupancy = dict(self._occupancy)
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
        return cloned

    # ---- node insertion and iteration ---------------------------------------

    def _clear_occupancy_for_node(self, node: Node) -> None:
        owner = node.key
        for cell in member_keys(node):
            if self._occupancy.get(cell) == owner:
                del self._occupancy[cell]

    def _register_occupancy(self, node: Node) -> None:
        owner = node.key
        members = member_keys(node)
        for cell in members:
            existing = self._occupancy.get(cell)
            if existing is not None and existing != owner:
                raise ValueError(f"Cell occupancy conflict: {cell} is already owned by {existing}")
        if owner in self._nodes:
            self._clear_occupancy_for_node(self._nodes[owner])
        for cell in members:
            self._occupancy[cell] = owner

    def add_node(self, node: Node) -> None:
        key = node.key
        self._register_occupancy(node)
        self._nodes[key] = node
        self._edges.setdefault(key, set())
        self._reverse_edges.setdefault(key, set())
        for hook in self._hooks:
            hook(key, node)

    def cell_owner(self, cell_key: NodeKey) -> NodeKey | None:
        """Return the owning node key for a single workbook cell, if any.

        Cell nodes own themselves. Multi-cell nodes own each expanded member.
        Raises `ValueError` when `cell_key` is not a single-cell address.
        """
        parsed = parse_node_key(cell_key)
        if not isinstance(parsed, CellKey):
            raise ValueError(f"Expected a single-cell key, got: {cell_key!r}")
        return self._occupancy.get(str(parsed))

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
        """Return an immutable `NodeView` snapshot, or `None` if missing.

        Lookup is by exact graph key. Member cells of a multi-cell node are not
        present as their own keys — use `locate_cell` / `cell_owner` first.
        """
        node = self._nodes.get(normalize_key(key))
        if node is None:
            return None
        return node_to_view(node)

    def get_dependencies(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return an immutable snapshot of `key`'s dependencies (cells it reads).

        Endpoints are returned exactly as stored (member cells are not rewritten).
        Use `get_dependency_nodes` when owner graph keys are required.
        """
        deps = self._edges.get(normalize_key(key))
        if not deps:
            return frozenset()
        return frozenset(deps)

    def get_dependents(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return an immutable snapshot of cells that depend on `key`.

        Endpoints are returned exactly as stored (member cells are not rewritten).
        """
        deps = self._reverse_edges.get(normalize_key(key))
        if not deps:
            return frozenset()
        return frozenset(deps)

    def resolve_endpoint(self, key: NodeKey) -> NodeKey | None:
        """Map an edge endpoint to a stored node key (exact or occupancy owner).

        Returns `None` when the key is neither a graph node nor an occupied
        member cell.
        """
        return self._resolve_graph_endpoint(key)

    def get_dependency_nodes(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return dependencies resolved to stored graph node keys.

        Member-cell endpoints map to their occupancy owners. Unresolvable
        dangling endpoints are omitted.
        """
        out: set[NodeKey] = set()
        for dep in self.get_dependencies(key):
            resolved = self.resolve_endpoint(dep)
            if resolved is not None:
                out.add(resolved)
        return frozenset(out)

    def get_edge_attrs(self, from_key: NodeKey, to_key: NodeKey) -> EdgeAttrs:
        """Return a typed snapshot of the attributes on edge `from_key -> to_key`.

        Lookup uses exact stored endpoints (no occupancy rewrite). When the edge
        does not exist, returns an `EdgeAttrs` with all fields set to `None`.
        """
        fk = normalize_key(from_key)
        tk = normalize_key(to_key)
        if tk not in self._edges.get(fk, set()):
            return EdgeAttrs()
        return EdgeAttrs(
            guard=self._guards.get((fk, tk)),
            provenance=self._edge_provenance.get((fk, tk)),
        )

    def get_edge_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None:
        """Return the guard on edge `from_key -> to_key`, or `None` if none.

        Lookup uses exact stored endpoints (no occupancy rewrite).
        """
        fk = normalize_key(from_key)
        tk = normalize_key(to_key)
        v = self._guards.get((fk, tk))
        return v if isinstance(v, GuardExpr) else None

    def is_guarded(self, from_key: NodeKey, to_key: NodeKey) -> bool:
        """Return whether edge `from_key -> to_key` carries a guard.

        Membership-only: does not retrieve the guard AST. Prefer this over
        `get_edge_guard(...) is not None` when the expression is unused.
        """
        fk = normalize_key(from_key)
        tk = normalize_key(to_key)
        return (fk, tk) in self._guards

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
        node.set_metadata(metadata)

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
        Node hooks are not invoked.

        Removing a multi-cell node clears occupancy for every expanded member.
        Removing a member cell while it is owned by a multi-cell node raises
        `ValueError`. Absent keys that are not occupied members are a no-op.
        """
        nk = normalize_key(key)
        if nk not in self._nodes:
            owner = self._occupancy.get(nk)
            if owner is not None and owner != nk:
                raise ValueError(
                    f"Cannot remove member cell {nk!r}; owned by multi-cell node {owner!r}"
                )
            return
        node = self._nodes[nk]
        for dep in list(self._edges.get(nk, set())):
            self._remove_edge(nk, dep)
        for dependent in list(self._reverse_edges.get(nk, set())):
            self._remove_edge(dependent, nk)
        self._clear_occupancy_for_node(node)
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
        """Iterate over (key, node) pairs for nodes that contain formulas.

        Formula cells are identified by `normalized_formula`, which is always
        stored; the raw `formula` string is opt-in (see `store_raw_formula`).
        """
        for key, node in self._nodes.items():
            if node.normalized_formula is not None:
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
            source=(k for k, node in self._nodes.items() if node.normalized_formula is not None),
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

    def _resolve_graph_endpoint(self, key: NodeKey) -> NodeKey | None:
        """Map an edge endpoint to a stored node key (exact or occupancy owner)."""
        nk = normalize_key(key)
        if nk in self._nodes:
            return nk
        # Member cell of a multi-cell node → owning graph key.
        owner = self._occupancy.get(nk)
        if owner is not None and owner in self._nodes:
            return owner
        return None

    def _unconditional_adjacency(self) -> dict[NodeKey, set[NodeKey]]:
        out: dict[NodeKey, set[NodeKey]] = {k: set() for k in self._nodes}
        for k in self._nodes:
            for dep in self._edges.get(k, ()):
                if self.is_guarded(k, dep):
                    continue
                resolved = self._resolve_graph_endpoint(dep)
                if resolved is not None:
                    out[k].add(resolved)
        return out

    def _all_adjacency(self) -> dict[NodeKey, set[NodeKey]]:
        out: dict[NodeKey, set[NodeKey]] = {k: set() for k in self._nodes}
        for k in self._nodes:
            for dep in self._edges.get(k, ()):
                resolved = self._resolve_graph_endpoint(dep)
                if resolved is not None:
                    out[k].add(resolved)
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
                    prov = self._edge_provenance.get((d_key, t_key))
                    if not compression_safe_provenance(prov):
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

        Collapses nodes when substitution is safe. Both identity-transit forwarding
        and formula inlining skip `is_target` nodes and any keys in `preserve`
        (external consumers such as series-bound public addresses). Forwarding
        targets are also protected from later inlining.

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
                if t_node is None or t_node.is_leaf or t_node.normalized_formula is None:
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
            return removed
        finally:
            clear_identity_singleton_ref_cache()

    # ---- serialization ------------------------------------------------------

    def __reduce_ex__(self, protocol: SupportsIndex, /) -> str | tuple[Any, ...]:
        """Pickle via a multipart blob so unpickle peak stays near final size."""
        del protocol
        return (loads_graph_blob, (dumps_graph_blob(self),))

    def __getstate__(self) -> dict[str, Any]:
        """Legacy compact state dict (kept for direct callers and old tests)."""
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
            "_edge_provenance": [
                (idx[a], idx[b], p) for (a, b), p in self._edge_provenance.items()
            ],
            "_hooks": self._hooks,
            "leaf_classification": self.leaf_classification,
            "sheet_order": list(self.sheet_order) if self.sheet_order is not None else None,
            "named_ranges": dict(self.named_ranges) if self.named_ranges else None,
            "named_range_ranges": (
                dict(self.named_range_ranges) if self.named_range_ranges else None
            ),
        }

    def __setstate__(self, state: dict[str, Any]) -> None:
        """Restore from a legacy v3 state dict, clearing intermediates as we go."""
        if not isinstance(state, dict) or state.get("v") != _PICKLE_VERSION:
            raise TypeError(
                "Unsupported or corrupted DependencyGraph pickle; rebuild the graph cache."
            )
        keys = state.pop("keys")
        key_index = {s: i for i, s in enumerate(keys)}

        nodes_raw = state.pop("_nodes")
        self._nodes = {keys[i]: n for i, n in nodes_raw.items()}
        nodes_raw.clear()

        edges_raw = state.pop("_edges")
        self._edges = {}
        for i, ds in edges_raw.items():
            self._edges[keys[i]] = {keys[d] for d in ds}
            ds.clear()
        edges_raw.clear()

        reverse_raw = state.pop("_reverse_edges")
        self._reverse_edges = {}
        for i, ds in reverse_raw.items():
            self._reverse_edges[keys[i]] = {keys[d] for d in ds}
            ds.clear()
        reverse_raw.clear()

        guards_raw = state.pop("_guards")
        self._guards = {
            (keys[a], keys[b]): _intern_guard_cell_refs(g, keys, key_index=key_index)
            for a, b, g in guards_raw
        }
        guards_raw.clear()

        provenance_raw = state.pop("_edge_provenance")
        self._edge_provenance = {(keys[a], keys[b]): p for a, b, p in provenance_raw}
        provenance_raw.clear()

        self._hooks = state.pop("_hooks")
        lc = state.pop("leaf_classification")
        if lc:
            self.leaf_classification = {keys[key_index[k]]: v for k, v in lc.items()}
        else:
            self.leaf_classification = None
        sheet_order = state.pop("sheet_order", None)
        self.sheet_order = list(sheet_order) if sheet_order else None
        nr = state.pop("named_ranges", None)
        self.named_ranges = dict(nr) if nr else None
        nrr = state.pop("named_range_ranges", None)
        self.named_range_ranges = dict(nrr) if nrr else None
        self.sheet_bounds = None
        self.preparsed_formulas = None
        state.clear()
        self._rebuild_occupancy()

    def _rebuild_occupancy(self) -> None:
        """Rebuild the cell→owner occupancy index from stored nodes."""
        self._occupancy = {}
        for node in self._nodes.values():
            owner = node.key
            for cell in member_keys(node):
                existing = self._occupancy.get(cell)
                if existing is not None and existing != owner:
                    raise ValueError(
                        f"Cell occupancy conflict while rebuilding: {cell} "
                        f"owned by both {existing} and {owner}"
                    )
                self._occupancy[cell] = owner

    # ---- internal edge mutation --------------------------------------------

    def _remove_edge(self, from_key: NodeKey, to_key: NodeKey) -> None:
        self._edges.setdefault(from_key, set()).discard(to_key)
        self._reverse_edges.setdefault(to_key, set()).discard(from_key)
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
            if isinstance(prov, EdgeProvenance) and prov.direct_sites_normalized and new_norm:
                new_norm = replace_substrings_at_spans(
                    new_norm, prov.direct_sites_normalized, r_key
                )
            elif new_norm and t_key in new_norm:
                new_norm = new_norm.replace(t_key, r_key)

            if record is not None and before_normalized != new_norm:
                record.formula_rewrites.append(
                    FormulaRewrite(
                        dependent=d_key,
                        before_normalized=before_normalized,
                        after_normalized=new_norm,
                    )
                )

            d_node.normalized_formula = new_norm

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
            direct_provenance_for_key_in_normalized,
            merge_inline_edge_guards,
            refresh_direct_sites,
            substitute_body_at_spans,
        )

        t_node = self._nodes.get(t_key)
        d_node = self._nodes.get(d_key)
        if t_node is None or d_node is None:
            return
        if t_node.normalized_formula is None:
            return

        prov = self._edge_provenance.get((d_key, t_key))
        if not isinstance(prov, EdgeProvenance):
            return

        before_normalized = d_node.normalized_formula
        new_norm = before_normalized
        if new_norm is not None and prov.direct_sites_normalized:
            new_norm = substitute_body_at_spans(
                new_norm,
                prov.direct_sites_normalized,
                t_node.normalized_formula,
            )

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

        d_node.normalized_formula = new_norm

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
    for a, b in g._edge_provenance:
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

    return intern_guard(rec(expr))


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

        for raw_w in graph._edges.get(v, ()):
            w = graph._resolve_graph_endpoint(raw_w)
            if w is None or w not in nodes:
                continue
            guard = graph._guards.get((v, raw_w))
            if guard is None:
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

        for raw_w in graph._edges.get(v, ()):
            w = graph._resolve_graph_endpoint(raw_w)
            if w is None or w not in nodes:
                continue
            guard = graph._guards.get((v, raw_w))
            if guard is None:
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
