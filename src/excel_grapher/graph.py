from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from typing import Any

from .node import Node, NodeKey

NodeHook = Callable[[NodeKey, Node], None]


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

    def add_edge(self, from_key: NodeKey, to_key: NodeKey, **attrs: Any) -> None:
        """Add edge: from_key depends on to_key (from_key -> to_key)."""
        self._edges.setdefault(from_key, set()).add(to_key)
        self._reverse_edges.setdefault(to_key, set()).add(from_key)
        if attrs:
            self._edge_attrs[(from_key, to_key)] = dict(attrs)
        else:
            self._edge_attrs.setdefault((from_key, to_key), {})

    def dependencies(self, key: NodeKey) -> set[NodeKey]:
        return self._edges.get(key, set())

    def dependents(self, key: NodeKey) -> set[NodeKey]:
        return self._reverse_edges.get(key, set())

    def edge_attrs(self, from_key: NodeKey, to_key: NodeKey) -> dict[str, Any]:
        return self._edge_attrs.get((from_key, to_key), {})

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

    def evaluation_order(self) -> list[NodeKey]:
        """
        Return nodes in dependency-first order (leaves before formulas that use them).

        Edge direction is A -> B meaning A depends on B. This method returns an
        ordering suitable for sequential evaluation (dependencies first).
        """
        order: list[NodeKey] = []
        perm: set[NodeKey] = set()
        temp: set[NodeKey] = set()

        def visit(n: NodeKey) -> None:
            if n in perm:
                return
            if n in temp:
                raise ValueError(f"Cycle detected involving {n}")
            temp.add(n)
            for dep in self.dependencies(n):
                if dep in self._nodes:
                    visit(dep)
            temp.remove(n)
            perm.add(n)
            order.append(n)

        for key in list(self._nodes.keys()):
            if key not in perm:
                visit(key)

        return order

