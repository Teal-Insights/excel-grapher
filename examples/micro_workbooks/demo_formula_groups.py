#!/usr/bin/env python3
"""Hand-built and coalesced formula-group nodes — Issues 1–3 demo.

Issue 1 (address model): multi-cell nodes keyed by `RangeKey` / `UnionKey`,
unique occupancy (no member cell nodes), `locate_cell` ownership lookup.

Issue 2 (evaluator + codegen): shared skeleton + per-member bindings,
`specialize_group`, lazy member eval, one `_group_*` helper in export.

Issue 3 (detection): `coalesce_formula_groups` / builder `formula_groups=True`
rewrite same-shape cell families into those groups.

Run from the repo root::

    uv run python examples/micro_workbooks/demo_formula_groups.py
"""

from __future__ import annotations

from excel_grapher.core.address_keys import (
    members_to_node_key,
    parse_node_key,
)
from excel_grapher.core.formula_ast import (
    AddressHoleNode,
    AddressLeafKind,
    BinaryOpNode,
    CellRefNode,
    NumberNode,
)
from excel_grapher.evaluator.errors import FormulaGroupKeyError
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.name_utils import address_to_python_name
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher.formula_groups import (
    coalesce_formula_groups,
    shape_fingerprint,
    specialize_group,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    locate_cell,
    make_cell_node,
    make_union_node,
    member_keys,
)


def _scale_skeleton() -> BinaryOpNode:
    """Shared template: `<CELL hole> * 10`."""
    return BinaryOpNode(
        op="*",
        left=AddressHoleNode(kind=AddressLeafKind.cell, slot=0),
        right=NumberNode(value=10.0),
    )


def build_row_stripe_group() -> tuple[DependencyGraph, str, tuple[str, ...]]:
    """Contiguous one-row formula group `Sheet1!D63:F63`."""
    members = ("Sheet1!D63", "Sheet1!E63", "Sheet1!F63")
    skeleton = _scale_skeleton()
    bindings = {
        "Sheet1!D63": (CellRefNode(address="Sheet1!D1"),),
        "Sheet1!E63": (CellRefNode(address="Sheet1!E1"),),
        "Sheet1!F63": (CellRefNode(address="Sheet1!F1"),),
    }
    group = make_union_node(
        members,
        is_leaf=False,
        shape_fingerprint=shape_fingerprint(skeleton),
        skeleton=skeleton,
        member_bindings=bindings,
        metadata={"role": "scaled-inputs"},
    )
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    g.add_node(make_cell_node("Sheet1", "D", 1, value=1.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "E", 1, value=2.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "F", 1, value=3.0, is_leaf=True))
    g.add_node(group)
    for leaf in ("Sheet1!D1", "Sheet1!E1", "Sheet1!F1"):
        g.add_edge(group.key, leaf)
    return g, group.key, members


def build_cross_sheet_union() -> tuple[DependencyGraph, str, tuple[str, ...]]:
    """Non-contiguous cross-sheet union with the same scale template."""
    members = ("Sheet1!D63", "Sheet2!B10")
    skeleton = _scale_skeleton()
    bindings = {
        "Sheet1!D63": (CellRefNode(address="Sheet1!D1"),),
        "Sheet2!B10": (CellRefNode(address="Sheet2!Z9"),),
    }
    group = make_union_node(
        members,
        is_leaf=False,
        shape_fingerprint=shape_fingerprint(skeleton),
        skeleton=skeleton,
        member_bindings=bindings,
        metadata={"role": "cross-sheet"},
    )
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    g.add_node(make_cell_node("Sheet1", "D", 1, value=4.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet2", "Z", 9, value=5.0, is_leaf=True))
    g.add_node(group)
    g.add_edge(group.key, "Sheet1!D1")
    g.add_edge(group.key, "Sheet2!Z9")
    return g, group.key, members


def demo_issue1_address_model() -> None:
    print("=" * 72)
    print("Issue 1 — Address model + storage + locate")
    print("=" * 72)
    print()

    print("Node key types (parse_node_key / members_to_node_key)")
    samples = (
        "Sheet1!E63",
        "Sheet1!D63:F63",
        "Sheet1!D63,Sheet2!B10",
    )
    for raw in samples:
        parsed = parse_node_key(raw)
        kind = type(parsed).__name__
        print(f"  {raw!r:32} -> {kind}")
    packed = members_to_node_key(["Sheet1!F63", "Sheet1!D63", "Sheet1!E63"])
    print(f"  members_to_node_key(F,D,E)     -> {packed!r} ({type(packed).__name__})")
    print()

    g, group_key, members = build_row_stripe_group()
    view = g.get_node(group_key)
    assert view is not None
    print("Hand-built contiguous group")
    print(f"  key:        {view.key}")
    print(f"  kind:       {view.kind}")
    print(f"  address:    {view.address!r} ({type(view.address).__name__})")
    print(f"  members:    {tuple(member_keys(view))}")
    print(f"  metadata:   {dict(view.metadata)}")
    print()

    print("Unique occupancy — member cells are not nodes")
    for m in members:
        print(f"  get_node({m!r}) -> {g.get_node(m)}")
        print(f"  cell_owner({m!r}) -> {g.cell_owner(m)!r}")
    print()

    print("locate_cell resolves members to the owning group")
    for probe in (*members, "Sheet1!D1", "Sheet1!Z99"):
        loc = locate_cell(g, probe)
        if loc is None:
            print(f"  {probe} -> (not in graph)")
        else:
            print(f"  {probe} -> {loc.kind} node {loc.node_key}")
    print()

    _, union_key, union_members = build_cross_sheet_union()
    print("Cross-sheet UnionKey (non-contiguous cover)")
    print(f"  key:     {union_key}")
    print(f"  parsed:  {type(parse_node_key(union_key)).__name__}")
    print(f"  members: {union_members}")
    print()


def demo_issue2_eval_codegen() -> None:
    print("=" * 72)
    print("Issue 2 — Template, specialize, evaluate, codegen")
    print("=" * 72)
    print()

    g, group_key, members = build_row_stripe_group()
    view = g.get_node(group_key)
    assert view is not None
    assert view.skeleton is not None
    assert view.member_bindings is not None

    print("Template fields on the group node")
    print(f"  shape_fingerprint: {view.shape_fingerprint}")
    print(f"  skeleton:          {view.skeleton!r}")
    print("  member_bindings:")
    for member, binds in view.member_bindings.items():
        print(f"    {member} -> {binds}")
    print()

    print("specialize_group fills holes for one member (shared by eval + codegen)")
    for member in members:
        specialized = specialize_group(view.skeleton, view.member_bindings[member])
        print(f"  {member}: {specialized!r}")
    print()

    print("FormulaEvaluator — public API is the member address")
    with FormulaEvaluator(g) as ev:
        for member in members:
            print(f"  evaluate({member!r}) -> {ev.evaluate(member)}")
        ev.clear_caches()
        _ = ev.evaluate("Sheet1!E63")
        print(f"  after evaluating E63 only, cache keys: {sorted(ev._cache)}")
        try:
            ev.evaluate(group_key)
        except FormulaGroupKeyError as exc:
            print(f"  evaluate(group_key) raises {type(exc).__name__}: {exc}")
    print()

    print("CodeGenerator — one _group_* helper + thin member wrappers")
    with CodeGenerator(g) as gen:
        code = gen.generate(targets=list(members))
        projected = gen._map_address_to_projected("Sheet1!E63")
    print(f"  map_to_projected(E63) -> address={projected.address!r}")
    print(f"                          parameters={projected.parameters}")
    print()
    # Print the interesting emitted defs only.
    interesting = [
        line
        for line in code.splitlines()
        if line.startswith("def _group_")
        or line.startswith("def cell_sheet1_")
        or line.startswith("    return _group_")
        or line.startswith("    '''Formula group")
    ]
    for line in interesting:
        print(f"  {line}")
    print()

    print("Exported wrappers match evaluator")
    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = ns["make_context"]
    assert callable(make_context)
    ctx = make_context()
    with FormulaEvaluator(g) as ev:
        for member in members:
            wrapper = ns[address_to_python_name(member)]
            assert callable(wrapper)
            exported = wrapper(ctx)
            evaluated = ev.evaluate(member)
            print(f"  {member}: export={exported}  eval={evaluated}")
    print()


def demo_issue3_coalesce() -> None:
    print("=" * 72)
    print("Issue 3 — Detect + coalesce (cell-only → group)")
    print("=" * 72)
    print()

    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet2", "A", 1, value=2.0, is_leaf=True))
    for member, leaf, sheet, col, formula in (
        ("Sheet1!B1", "Sheet1!A1", "Sheet1", "B", "=Sheet1!A1*10"),
        ("Sheet2!B1", "Sheet2!A1", "Sheet2", "B", "=Sheet2!A1*10"),
    ):
        g.add_node(
            make_cell_node(
                sheet,
                col,
                1,
                formula=formula,
                normalized_formula=formula,
                is_leaf=False,
                is_target=True,
            )
        )
        g.add_edge(member, leaf)

    print("Before coalesce (cell-only)")
    print(f"  nodes: {g.keys(order='workbook')}")
    print(f"  target_keys: {g.target_keys()}")
    print()

    report = coalesce_formula_groups(g)
    print("After coalesce_formula_groups")
    print(f"  created_groups: {report.created_groups}")
    print(f"  skipped: {[(s.reason, s.members) for s in report.skipped_families]}")
    print(f"  nodes: {g.keys(order='workbook')}")
    print(f"  target_keys (still member addresses): {g.target_keys()}")
    for member in ("Sheet1!B1", "Sheet2!B1"):
        loc = locate_cell(g, member)
        print(f"  locate_cell({member}) -> {None if loc is None else loc.node_key}")
    with FormulaEvaluator(g) as ev:
        print(f"  evaluate(Sheet1!B1) -> {ev.evaluate('Sheet1!B1')}")
        print(f"  evaluate(Sheet2!B1) -> {ev.evaluate('Sheet2!B1')}")
    print()


def main() -> None:
    print()
    print("Formula-group nodes demo (Issues #391 / #392 / #393)")
    print("Unique occupancy: members are addresses, not cell nodes.")
    print()
    demo_issue1_address_model()
    demo_issue2_eval_codegen()
    demo_issue3_coalesce()


if __name__ == "__main__":
    main()
