"""Shared fixtures for compression unit tests."""

from __future__ import annotations

from collections.abc import Mapping

from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.formula_ast import AstNode, parse
from excel_grapher.core.types import CellValue
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node


def make_node(
    address: str,
    formula: str | None,
    value: CellValue = None,
    *,
    normalized_formula: str | None = None,
    is_leaf: bool | None = None,
) -> Node:
    """Build a graph node from a sheet-qualified address."""
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    if is_leaf is None:
        is_leaf = formula is None
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=normalized_formula if normalized_formula is not None else formula,
        value=value,
        is_leaf=is_leaf,
    )


def make_graph(*nodes: Node) -> DependencyGraph:
    """Build a dependency graph from explicit nodes."""
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def parse_formula(formula: str) -> AstNode:
    """Parse a normalized formula string into an AST."""
    return parse(formula)


def leaf_values(values: Mapping[str, CellValue]) -> list[Node]:
    """Build value-only leaf nodes from an address-to-value map."""
    return [make_node(address, None, value) for address, value in values.items()]
