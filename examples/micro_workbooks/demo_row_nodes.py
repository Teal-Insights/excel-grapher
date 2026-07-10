#!/usr/bin/env python3
"""Hand-built DependencyGraph demos for first-class row nodes.

Part 1 (issue #374): storage + edges + locate helpers on a mixed graph.
Part 2 (issue #377): executable stripe — evaluate members, show lazy
caching, and emit a parameterized `_row_*` helper via codegen.

`create_dependency_graph` still expands ranges to cell precedents; these graphs
are constructed by hand.

Run from the repo root::

    uv run python examples/micro_workbooks/demo_row_nodes.py
"""

from __future__ import annotations

from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator
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


def build_storage_demo_graph() -> DependencyGraph:
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


def build_demo_graph() -> DependencyGraph:
    """Build one template row node with no member cell nodes.

    Template `=Sheet1!D35*2` with `varying_ref_slots=(0,)` specializes per column
    so `evaluate("Sheet1!E63")` rewrites the varying ref to `E35` and returns 10.
    Span is `D63:F63` so `evaluate_row` can show a subrange vs the full stripe.
    """
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    leaves = (("D", 3), ("E", 5), ("F", 7))
    for col, value in leaves:
        g.add_node(_leaf("Sheet1", col, 35, value))
    row = make_row_node(
        "Sheet1",
        63,
        "D",
        "F",
        formula="=Sheet1!D35*2",
        normalized_formula="=Sheet1!D35*2",
        varying_ref_slots=(0,),
        is_leaf=False,
        is_target=True,
        metadata={"role": "demo-template"},
    )
    g.add_node(row)
    for col, _value in leaves:
        g.add_edge(row.key, f"Sheet1!{col}35")
    return g


def demo_storage_and_locate() -> None:
    g = build_storage_demo_graph()
    row_key = "Sheet1!D63:Y63"

    view = g.get_node(row_key)
    assert view is not None
    print("=== Part 1: storage + locate ===")
    print("Row node")
    print(f"  key:      {view.key}")
    print(f"  kind:     {view.kind}")
    print(f"  address:  {view.address}")
    print(f"  extent:   {view.min_col}{view.row}:{view.max_col}{view.row}")
    print(f"  metadata: {dict(view.metadata)}")
    print()

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
    print()


def demo_eval_and_codegen() -> None:
    g = build_demo_graph()
    row_key = "Sheet1!D63:F63"
    row = g.get_node(row_key)
    assert row is not None

    print("=== Part 2: eval + codegen ===")
    print("Row template")
    print(f"  key:                 {row.key}")
    print(f"  normalized_formula:  {row.normalized_formula}")
    print(f"  varying_ref_slots:   {row.varying_ref_slots}")
    print(f"  member cells:        {row_member_keys(row)}")
    print(f"  cell nodes at D63/E63: {g.get_node('Sheet1!D63')}, {g.get_node('Sheet1!E63')}")
    print()

    with FormulaEvaluator(g) as ev:
        print("evaluate(member) via locate_cell + specialize")
        for member in ("Sheet1!D63", "Sheet1!E63", "Sheet1!F63"):
            print(f"  {member} -> {ev.evaluate(member)}")
        print(f"  cache after all members: {sorted(ev._cache)}")
        print()

        ev.clear_caches()
        print("Laziness: evaluate E63 only")
        print(f"  E63 -> {ev.evaluate('Sheet1!E63')}")
        print(f"  cache keys: {sorted(ev._cache)}")
        print(f"  D63 cached? {'Sheet1!D63' in ev._cache}")
        print()

        try:
            ev.evaluate(row_key)
        except ValueError as exc:
            print(f"Row-key evaluate() rejected: {exc}")
        print()

        ev.clear_caches()
        print("evaluate_row (full stripe vs subrange)")
        print(f"  evaluate_row({row_key!r}) -> {ev.evaluate_row(row_key)}")
        ev.clear_caches()
        print(f"  evaluate_row('Sheet1!E63:F63') -> {ev.evaluate_row('Sheet1!E63:F63')}")
        print(f"  cache after subrange: {sorted(ev._cache)}")
        print(f"  D63 cached? {'Sheet1!D63' in ev._cache}")
        print()

    code = CodeGenerator(g).generate(["Sheet1!D63", "Sheet1!E63", "Sheet1!F63"])
    print("Codegen excerpt (helpers + wrappers)")
    for line in code.splitlines():
        stripped = line.lstrip()
        if (
            stripped.startswith("def _row_")
            or stripped.startswith("def cell_sheet1_")
            or stripped.startswith("return _row_")
            or 'f"Sheet1!{column}' in stripped
        ):
            print(f"  {line}")
    print()

    ns: dict[str, object] = {}
    exec(code, ns)
    compute_all = ns["compute_all"]
    assert callable(compute_all)
    print(f"compute_all() -> {compute_all()}")


def main() -> None:
    demo_storage_and_locate()
    demo_eval_and_codegen()


if __name__ == "__main__":
    main()
