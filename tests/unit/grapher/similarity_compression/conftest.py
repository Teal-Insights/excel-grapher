"""Shared graph builders for similarity-compression unit tests."""

from __future__ import annotations

from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node


def make_node(
    key: str,
    formula: str | None,
    normalized: str | None,
    *,
    is_leaf: bool = False,
    is_target: bool = False,
) -> Node:
    sheet, rest = key.split("!", 1)
    if sheet.startswith("'"):
        sheet = sheet[1:-1]
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=normalized,
        value=None,
        is_leaf=is_leaf,
        is_target=is_target,
    )


def direct_edge(
    graph: DependencyGraph,
    dependent: str,
    precedent: str,
    *,
    formula: str | None = None,
    normalized: str | None = None,
) -> None:
    dr = DependencyCause.direct_ref
    dep_node = graph.get_node(dependent)
    assert dep_node is not None
    f = formula if formula is not None else dep_node.formula
    n = normalized if normalized is not None else dep_node.normalized_formula
    assert f is not None and n is not None
    ref = precedent
    i_f = f.index(ref)
    i_n = n.index(ref)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=((i_f, i_f + len(ref)),),
            direct_sites_normalized=((i_n, i_n + len(ref)),),
        ),
    )
