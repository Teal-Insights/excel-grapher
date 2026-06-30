"""Materialize TACO index queries to cell-level keys."""

from __future__ import annotations

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import parse_address
from excel_grapher.grapher.node import NodeKey

from .index import TacoIndex
from .patterns import materialize_dependents_for_edge, materialize_precedents_for_edge


def materialize_dependents(index: TacoIndex, query: NodeKey) -> set[NodeKey]:
    """Expand dependent relationships for one cell to sheet-qualified keys."""
    out: set[NodeKey] = set()
    sheet, col, row = _split_key(query)
    for edge_index in index._prec_spatial.query_point(sheet, col, row):
        edge = index.compressed_edges[edge_index]
        out.update(
            materialize_dependents_for_edge(edge.precedent, edge.dependent, edge.meta, query)
        )
    for dep_key in index._single_prec.get(query, []):
        out.add(dep_key)
    return out


def materialize_precedents(index: TacoIndex, query: NodeKey) -> set[NodeKey]:
    """Expand precedent relationships for one cell to sheet-qualified keys."""
    out: set[NodeKey] = set()
    sheet, col, row = _split_key(query)
    for edge_index in index._dep_spatial.query_point(sheet, col, row):
        edge = index.compressed_edges[edge_index]
        out.update(
            materialize_precedents_for_edge(edge.precedent, edge.dependent, edge.meta, query)
        )
    for prec_key in index._single_dep.get(query, []):
        out.add(prec_key)
    return out


def _split_key(key: NodeKey) -> tuple[str, str, int]:
    sheet, coord = parse_address(key)
    col, row = fastpyxl.utils.cell.coordinate_from_string(coord)
    return sheet, col, row
