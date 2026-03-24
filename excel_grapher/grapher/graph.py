from __future__ import annotations

import warnings
from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from typing import Any

from .dependency_provenance import EdgeProvenance, merge_edge_provenance
from .guard import GuardConstraints, GuardExpr, Or, or_guard
from .node import Node, NodeKey

NodeHook = Callable[[NodeKey, Node], None]


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


@dataclass
class DependencyGraph:
    _nodes: dict[NodeKey, Node] = field(default_factory=dict)
    _edges: dict[NodeKey, set[NodeKey]] = field(default_factory=dict)  # node -> deps
    _reverse_edges: dict[NodeKey, set[NodeKey]] = field(default_factory=dict)  # node -> dependents
    _edge_attrs: dict[tuple[NodeKey, NodeKey], dict[str, Any]] = field(default_factory=dict)
    _hooks: list[NodeHook] = field(default_factory=list)

    def add_node(self, node: Node) -> None:
        key = node.key
        self._nodes[key] = node
        self._edges.setdefault(key, set())
        self._reverse_edges.setdefault(key, set())
        for hook in self._hooks:
            hook(key, node)

    def get_node(self, key: NodeKey) -> Node | None:
        return self._nodes.get(key)

    def __contains__(self, key: NodeKey) -> bool:
        return key in self._nodes

    def __iter__(self) -> Iterator[NodeKey]:
        return iter(self._nodes)

    def __len__(self) -> int:
        return len(self._nodes)

    def add_edge(
        self,
        from_key: NodeKey,
        to_key: NodeKey,
        *,
        guard: GuardExpr | None = None,
        **attrs: Any,
    ) -> None:
        """Add edge: from_key depends on to_key (from_key -> to_key)."""
        self._edges.setdefault(from_key, set()).add(to_key)
        self._reverse_edges.setdefault(to_key, set()).add(from_key)
        existing = self._edge_attrs.get((from_key, to_key))
        if existing is None:
            merged_guard = guard
            merged: dict[str, Any] = {}
        else:
            # Distinguish "missing" from explicit unconditional guard=None.
            had_guard = "guard" in existing
            existing_guard = existing.get("guard")

            if not had_guard:
                merged_guard = guard
            else:
                # Unconditional dominates any guarded variant.
                if existing_guard is None or guard is None:
                    merged_guard = None
                elif existing_guard == guard:
                    merged_guard = guard
                else:
                    merged_guard = or_guard(existing_guard, guard)

            merged = dict(existing)

        merged.update(attrs)
        merged["guard"] = merged_guard
        prov_new = attrs.get("provenance")
        if prov_new is not None and isinstance(prov_new, EdgeProvenance):
            old_prov = merged.get("provenance")
            merged["provenance"] = merge_edge_provenance(
                old_prov if isinstance(old_prov, EdgeProvenance) else None,
                prov_new,
            )
        self._edge_attrs[(from_key, to_key)] = merged

    def dependencies(self, key: NodeKey) -> set[NodeKey]:
        return self._edges.get(key, set())

    def dependents(self, key: NodeKey) -> set[NodeKey]:
        return self._reverse_edges.get(key, set())

    def edge_attrs(self, from_key: NodeKey, to_key: NodeKey) -> dict[str, Any]:
        return self._edge_attrs.get((from_key, to_key), {})

    def edge_guard(self, from_key: NodeKey, to_key: NodeKey) -> GuardExpr | None:
        v = self._edge_attrs.get((from_key, to_key), {}).get("guard")
        return v if isinstance(v, GuardExpr) else None

    def register_hook(self, hook: NodeHook) -> None:
        self._hooks.append(hook)

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
        """Iterate over (key, node) pairs for leaf nodes (no formula)."""
        for key, node in self._nodes.items():
            if node.is_leaf:
                yield key, node

    def formula_keys(self) -> list[NodeKey]:
        """Return sorted list of keys for nodes that contain formulas."""
        return sorted(k for k, node in self._nodes.items() if node.formula is not None)

    def leaf_keys(self) -> list[NodeKey]:
        """Return sorted list of keys for leaf nodes."""
        return sorted(k for k, node in self._nodes.items() if node.is_leaf)

    def roots(self) -> Iterator[NodeKey]:
        for key in self._nodes:
            if not self._reverse_edges.get(key):
                yield key

    def _unconditional_adjacency(self) -> dict[NodeKey, set[NodeKey]]:
        out: dict[NodeKey, set[NodeKey]] = {k: set() for k in self._nodes}
        for k in self._nodes:
            for dep in self.dependencies(k):
                if dep not in self._nodes:
                    continue
                if self.edge_guard(k, dep) is None:
                    out[k].add(dep)
        return out

    def _all_adjacency(self) -> dict[NodeKey, set[NodeKey]]:
        out: dict[NodeKey, set[NodeKey]] = {k: set() for k in self._nodes}
        for k in self._nodes:
            for dep in self.dependencies(k):
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

    def evaluation_order(
        self, *, strict: bool = True, iterate_enabled: bool | None = None
    ) -> list[NodeKey]:
        """
        Return nodes in dependency-first order (leaves before formulas that use them).

        Edge direction is A -> B meaning A depends on B. This method returns an
        ordering suitable for sequential evaluation (dependencies first).

        If ``iterate_enabled`` is True (workbook has iterative calculation on), any
        must-cycle or may-cycle is rejected: generated Python does not emulate Excel's
        iterative convergence. Pass ``False`` or ``None`` to apply the usual strict /
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
            for dep in adjacency.get(n, set()):
                if dep in exclude:
                    continue
                if dep in self._nodes and dep not in exclude:
                    visit(dep)
            temp.remove(n)
            perm.add(n)
            order.append(n)

        for key in list(self._nodes.keys()):
            if key in exclude:
                continue
            if key not in perm:
                visit(key)

        return order

    def compress_identity_transits(self) -> list[NodeKey]:
        """
        Remove identity transit nodes (formula is a single cell reference to one dependency),
        rewrite dependents' formulas, and rewire edges. Requires dependency provenance from
        graph construction with ``capture_dependency_provenance=True`` for safe edges.

        Node hooks are not invoked for removed or updated nodes.

        Returns:
            Keys of removed transit nodes, in removal order.
        """
        from .compression import (
            compression_safe_provenance,
            direct_provenance_for_key_in_strings,
            is_identity_transit,
            replace_substrings_at_spans,
        )

        removed: list[NodeKey] = []
        while True:
            found: tuple[NodeKey, NodeKey] | None = None
            for t_key in sorted(self._nodes.keys()):
                r_key = is_identity_transit(self, t_key)
                if r_key is None:
                    continue
                if not self.dependents(t_key):
                    continue
                ok = True
                for d_key in self.dependents(t_key):
                    prov = self.edge_attrs(d_key, t_key).get("provenance")
                    if not compression_safe_provenance(
                        prov if isinstance(prov, EdgeProvenance) else None
                    ):
                        ok = False
                        break
                if not ok:
                    continue
                found = (t_key, r_key)
                break

            if found is None:
                break

            t_key, r_key = found
            if t_key in self._nodes and is_identity_transit(self, t_key) == r_key:
                self._compress_one_transit(t_key, r_key)
                removed.append(t_key)

        return removed

    def _remove_edge(self, from_key: NodeKey, to_key: NodeKey) -> None:
        self._edges.setdefault(from_key, set()).discard(to_key)
        self._reverse_edges.setdefault(to_key, set()).discard(from_key)
        self._edge_attrs.pop((from_key, to_key), None)

    def _compress_one_transit(self, t_key: NodeKey, r_key: NodeKey) -> None:
        from .compression import (
            direct_provenance_for_key_in_strings,
            replace_substrings_at_spans,
        )

        for d_key in list(self.dependents(t_key)):
            attrs = self.edge_attrs(d_key, t_key)
            prov = attrs.get("provenance")
            guard = self.edge_guard(d_key, t_key)
            d_node = self.get_node(d_key)
            if d_node is None:
                continue

            new_formula = d_node.formula
            new_norm = d_node.normalized_formula
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

            object.__setattr__(d_node, "formula", new_formula)
            object.__setattr__(d_node, "normalized_formula", new_norm)

            self._remove_edge(d_key, t_key)
            new_prov = direct_provenance_for_key_in_strings(new_formula, new_norm, r_key)
            self.add_edge(d_key, r_key, guard=guard, provenance=new_prov)

        for dep in list(self.dependencies(t_key)):
            self._remove_edge(t_key, dep)
        self._nodes.pop(t_key, None)
        self._edges.pop(t_key, None)
        self._reverse_edges.pop(t_key, None)


def _scc_cycles(adj: dict[NodeKey, set[NodeKey]]) -> list[set[NodeKey]]:
    """
    Return SCCs that are cyclic (size>1 or self-loop).
    """
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
    """
    Find one cycle path within the given node subset (best-effort).
    """
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
    """
    Conjoin an edge guard onto the current constraints, returning the resulting
    constraint sets.

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
    """
    Return True if there exists at least one cycle within `nodes` whose accumulated
    edge guards are jointly consistent (symbolic, no evaluation).
    """
    visited: set[tuple[NodeKey, GuardConstraints]] = set()
    on_stack: set[NodeKey] = set()

    def dfs(v: NodeKey, c: GuardConstraints) -> bool:
        state = (v, c)
        if state in visited:
            return False
        visited.add(state)
        on_stack.add(v)

        for w in graph.dependencies(v):
            if w not in nodes:
                continue
            guard = graph.edge_guard(v, w)
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
    """
    Best-effort: find one feasible cycle path within `nodes` (symbolic constraints).
    """
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

        for w in graph.dependencies(v):
            if w not in nodes:
                continue
            guard = graph.edge_guard(v, w)
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

