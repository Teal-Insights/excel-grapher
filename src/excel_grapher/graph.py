from __future__ import annotations

import warnings
from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from typing import Any

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
        for key, node in self._nodes.items():
            if node.is_leaf:
                yield key

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

    def evaluation_order(self, *, strict: bool = True) -> list[NodeKey]:
        """
        Return nodes in dependency-first order (leaves before formulas that use them).

        Edge direction is A -> B meaning A depends on B. This method returns an
        ordering suitable for sequential evaluation (dependencies first).
        """
        report = self.cycle_report()
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

