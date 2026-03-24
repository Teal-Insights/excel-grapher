from __future__ import annotations

from excel_grapher.core.formula_ast import (
    CellRefNode,
    FormulaParseError,
    UnaryOpNode,
    parse,
)

from .dependency_provenance import DependencyCause, EdgeProvenance
from .graph import DependencyGraph
from .node import NodeKey


def is_identity_transit(graph: DependencyGraph, transit_key: NodeKey) -> NodeKey | None:
    """
    If ``transit_key`` is a pure identity reference to exactly one dependency, return that
    dependency's key; otherwise return None.
    """
    node = graph.get_node(transit_key)
    if node is None or node.is_leaf or not node.normalized_formula:
        return None
    deps = graph.dependencies(transit_key)
    if len(deps) != 1:
        return None
    r_key = next(iter(deps))
    if graph.edge_guard(transit_key, r_key) is not None:
        return None
    try:
        ast = parse(node.normalized_formula)
    except FormulaParseError:
        return None
    while isinstance(ast, UnaryOpNode) and ast.op == "+":
        ast = ast.operand
    if not isinstance(ast, CellRefNode):
        return None
    r_node = graph.get_node(r_key)
    if r_node is None:
        return None
    if ast.address != r_node.key:
        return None
    return r_key


def replace_substrings_at_spans(formula: str, spans: tuple[tuple[int, int], ...], replacement: str) -> str:
    """Replace each ``[a,b)`` span in ``formula`` with ``replacement`` (right-to-left)."""
    out = formula
    for a, b in sorted(spans, reverse=True):
        if 0 <= a <= b <= len(out):
            out = out[:a] + replacement + out[b:]
    return out


def direct_provenance_for_key_in_strings(
    formula: str | None,
    normalized: str | None,
    dep_key: str,
) -> EdgeProvenance:
    """Build minimal direct-ref provenance by locating ``dep_key`` substrings."""
    sites_f: list[tuple[int, int]] = []
    sites_n: list[tuple[int, int]] = []
    if formula:
        sites_f.extend(_find_literal_spans(formula, dep_key))
    if normalized:
        sites_n.extend(_find_literal_spans(normalized, dep_key))
    return EdgeProvenance(
        causes=frozenset({DependencyCause.direct_ref}),
        direct_sites_formula=tuple(sites_f),
        direct_sites_normalized=tuple(sites_n),
    )


def _find_literal_spans(s: str, needle: str) -> list[tuple[int, int]]:
    if not needle:
        return []
    out: list[tuple[int, int]] = []
    i = 0
    while True:
        j = s.find(needle, i)
        if j < 0:
            break
        out.append((j, j + len(needle)))
        i = j + len(needle)
    return out


def compression_safe_provenance(prov: EdgeProvenance | None) -> bool:
    if prov is None:
        return False
    if DependencyCause.direct_ref not in prov.causes:
        return False
    unsafe = {
        DependencyCause.static_range,
        DependencyCause.dynamic_offset,
        DependencyCause.dynamic_indirect,
    }
    return not (prov.causes & unsafe)
