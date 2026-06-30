"""Materialize TACO index queries to cell-level keys."""

from __future__ import annotations

from excel_grapher.grapher.node import NodeKey

from .index import TacoIndex
from .patterns import rr_materialize_dependent, rr_materialize_precedent
from .types import PatternKind


def materialize_dependents(index: TacoIndex, query: NodeKey) -> set[NodeKey]:
    """Expand dependent relationships for one cell to sheet-qualified keys."""
    out: set[NodeKey] = set()
    for edge in index.compressed_edges:
        if edge.precedent.contains(query):
            if edge.meta.kind == PatternKind.rr:
                out.add(rr_materialize_dependent(edge.precedent, edge.dependent, query))
            else:
                out.update(edge.dependent.cell_keys())
    for single in index.single_edges:
        if single.precedent == query:
            out.add(single.dependent)
    return out


def materialize_precedents(index: TacoIndex, query: NodeKey) -> set[NodeKey]:
    """Expand precedent relationships for one cell to sheet-qualified keys."""
    out: set[NodeKey] = set()
    for edge in index.compressed_edges:
        if edge.dependent.contains(query):
            if edge.meta.kind == PatternKind.rr:
                out.add(rr_materialize_precedent(edge.dependent, edge.precedent, query))
            else:
                out.update(edge.precedent.cell_keys())
    for single in index.single_edges:
        if single.dependent == query:
            out.add(single.precedent)
    return out
