"""Helpers for tightening may-cycle feasibility (#533)."""

from __future__ import annotations

from collections.abc import Mapping

from excel_grapher.core.formula_ast import CellRefNode, resolve_cell_ref

from .node import Node, NodeKey


def identity_alias_map(nodes: Mapping[NodeKey, Node]) -> dict[NodeKey, NodeKey]:
    """Map identity-formula cells to the cell they copy, followed to a root.

    A node is an identity alias when `formula_ast` is a single `CellRefNode`.
    Relative axes resolve against the host cell. Cycles in the alias graph stop
    at the first repeat (the repeating node is the root).
    """
    parent: dict[NodeKey, NodeKey] = {}
    for key, node in nodes.items():
        ast = node.formula_ast
        if not isinstance(ast, CellRefNode):
            continue
        try:
            target = resolve_cell_ref(ast, key)
        except ValueError:
            continue
        parent[key] = target

    def root(key: NodeKey) -> NodeKey:
        seen: set[NodeKey] = set()
        while key in parent and key not in seen:
            seen.add(key)
            key = parent[key]
        return key

    return {key: root(key) for key in parent if root(key) != key}
