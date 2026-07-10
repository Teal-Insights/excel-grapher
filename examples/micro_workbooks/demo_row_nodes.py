#!/usr/bin/env python3
"""Hand-built DependencyGraph with first-class row nodes (issue #374).

`create_dependency_graph` still expands ranges to cell precedents. This script
shows the storage/edges API: construct a one-row node, wire cell↔row edges,
look up non-canonical keys, and print evaluation order + Mermaid.

Run from the repo root::

    uv run python examples/micro_workbooks/demo_row_nodes.py
"""

from __future__ import annotations

from excel_grapher.grapher.export import to_mermaid
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    locate_cell,
    locate_range,
    make_row_node,
    row_member_keys,
)


def _leaf(sheet: str, col: str, row: int, value: object = 0) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=None,
        normalized_formula=None,
        value=value,
        is_leaf=True,
    )


def _formula(sheet: str, col: str, row: int, formula: str) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=None,
        is_leaf=False,
        is_target=True,
    )


def build_demo_graph() -> DependencyGraph:
    """Build a tiny mixed graph: total cell depends on a one-row input stripe."""
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]

    # One-row precedent spanning D63:Y63 (hand-inserted; not from the builder).
    inputs = make_row_node("Sheet1", 63, "D", "Y", metadata={"role": "inputs"})
    total = _formula("Sheet1", "A", 63, "=SUM(D63:Y63)")
    flag = _leaf("Sheet1", "C", 1, value=True)

    g.add_node(inputs)
    g.add_node(total)
    g.add_node(flag)
    g.add_edge(total.key, inputs.key)  # cell → row
    g.add_edge(total.key, flag.key)  # cell → cell

    return g


def main() -> None:
    g = build_demo_graph()
    row_key = "Sheet1!D63:Y63"

    view = g.get_node(row_key)
    assert view is not None
    print("Row node")
    print(f"  key:      {view.key}")
    print(f"  kind:     {view.kind}")
    print(f"  address:  {view.address}")
    print(f"  extent:   {view.min_col}{view.row}:{view.max_col}{view.row}")
    print(f"  metadata: {dict(view.metadata)}")
    print()

    # Non-canonical spellings resolve to the same node.
    aliases = [
        "Sheet1!Y63:D63",
        "Sheet1!D63:Sheet1!Y63",
        "Sheet1!$D$63:$Y$63",
    ]
    print("Canonical lookups")
    for alias in aliases:
        found = g.get_node(alias)
        print(f"  {alias!r:30} -> {None if found is None else found.key}")
    print()

    members = row_member_keys(view)
    print(f"Member cell keys ({len(members)}): {members[0]} … {members[-1]}")
    print()

    print("Where does this cell live?")
    for probe in ("Sheet1!E63", "Sheet1!A63", "Sheet1!C1", "Sheet1!Z99"):
        loc = locate_cell(g, probe)
        if loc is None:
            print(f"  {probe} -> (not in graph)")
        else:
            print(f"  {probe} -> {loc.kind} node {loc.node_key} (column {loc.column})")
    print()

    print("Where does this subrange live?")
    for probe in ("Sheet1!E63:G63", "Sheet1!G63:E63", "Sheet1!D63:Y63", "Sheet1!C63:E63"):
        loc = locate_range(g, probe)
        if loc is None:
            print(f"  {probe} -> (not in graph)")
        else:
            print(
                f"  {probe} -> {loc.kind} node {loc.node_key} "
                f"(query {loc.min_col}{loc.row}:{loc.max_col}{loc.row})"
            )
    print()

    print("Edges")
    for key in g.keys(order="workbook"):
        deps = sorted(g.get_dependencies(key))
        if deps:
            print(f"  {key} depends on {deps}")
    print(f"  dependents of row: {sorted(g.get_dependents(row_key))}")
    print()

    print("Evaluation order")
    for key in g.evaluation_order():
        node = g.get_node(key)
        kind = node.kind if node is not None else NodeKind.cell
        print(f"  [{kind}] {key}")
    print()

    print("Mermaid")
    print(to_mermaid(g))


if __name__ == "__main__":
    main()
