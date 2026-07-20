#!/usr/bin/env python3
"""Workbook → graph → coalesce → eval / codegen for mixed formula groups.

Writes a spreadsheet with:

- a **row** of same-shape formulas (coalesce → `Sheet1!B2:D2`, `NodeKind.row`)
- a **column** of same-shape formulas (coalesce → `Sheet1!B3:B5`, column shape)
- **individual** formula cells that stay as cells (`below_min_size` / unique shape)

Then builds with `create_dependency_graph`, runs `coalesce_formula_groups`,
and checks that evaluator and codegen agree with and without coalescing on
**all** formula targets (row members, column members, and leftover cells).

Run from the repo root::

    uv run python examples/micro_workbooks/demo_formula_groups.py
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import xlsxwriter

from excel_grapher.core.address_keys import NodeShape
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.name_utils import address_to_python_name
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.export import to_mermaid
from excel_grapher.grapher.formula_groups import coalesce_formula_groups, specialize_group
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKind, locate_cell, member_keys

WORKBOOK = Path(__file__).with_name("demo_formula_groups.xlsx")

# Contiguous row family: =leaf*10 → Sheet1!B2:D2
ROW_MEMBERS = ("Sheet1!B2", "Sheet1!C2", "Sheet1!D2")
# Contiguous column family: =leaf+5 → Sheet1!B3:B5
COL_MEMBERS = ("Sheet1!B3", "Sheet1!B4", "Sheet1!B5")
# Stay as individual cells after coalesce.
CELL_TARGETS = ("Sheet1!F2", "Sheet1!G2", "Sheet1!H2")

TARGETS = (*ROW_MEMBERS, *COL_MEMBERS, *CELL_TARGETS)


def write_demo_workbook(path: Path) -> Path:
    """Write inputs + row / column / individual formula patterns."""
    path = Path(path)
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")

    # Row inputs B1:D1 and column inputs A3:A5; lone input F1.
    ws.write_number("B1", 1.0)
    ws.write_number("C1", 2.0)
    ws.write_number("D1", 3.0)
    ws.write_number("A3", 10.0)
    ws.write_number("A4", 20.0)
    ws.write_number("A5", 30.0)
    ws.write_number("F1", 7.0)

    # Row formula group (fingerprint *10).
    ws.write_formula("B2", "=Sheet1!B1*10", None, 10.0)
    ws.write_formula("C2", "=Sheet1!C1*10", None, 20.0)
    ws.write_formula("D2", "=Sheet1!D1*10", None, 30.0)

    # Column formula group (different fingerprint +5 so it does not merge with the row).
    ws.write_formula("B3", "=Sheet1!A3+5", None, 15.0)
    ws.write_formula("B4", "=Sheet1!A4+5", None, 25.0)
    ws.write_formula("B5", "=Sheet1!A5+5", None, 35.0)

    # Individual formula cells (unique / singleton families stay as cells).
    ws.write_formula("F2", "=Sheet1!F1*2", None, 14.0)
    ws.write_formula("G2", "=ABS(Sheet1!F1)", None, 7.0)
    ws.write_formula("H2", "=SUM(B2:D2)", None, 60.0)

    wb.close()
    return path


def build_cell_only_graph(workbook: Path) -> DependencyGraph:
    return create_dependency_graph(
        workbook,
        list(TARGETS),
        load_values=True,
        use_cached_dynamic_refs=True,
        formula_groups=False,
    )


def evaluate_targets(graph: DependencyGraph, targets: tuple[str, ...]) -> dict[str, object]:
    with FormulaEvaluator(graph) as ev:
        return {t: ev.evaluate(t) for t in targets}


def codegen_targets(
    graph: DependencyGraph, targets: tuple[str, ...]
) -> tuple[dict[str, object], str]:
    code = CodeGenerator(graph).generate(targets=list(targets))
    ns: dict[str, object] = {}
    exec(code, ns)
    compute_all = ns["compute_all"]
    assert callable(compute_all)
    results = cast(dict[str, object], compute_all())
    if all(t in results for t in targets):
        return {t: results[t] for t in targets}, code
    by_wrapper = {
        t: cast(Any, ns[address_to_python_name(t)])(cast(Any, ns["make_context"])())
        for t in targets
    }
    return by_wrapper, code


def _assert_maps_equal(
    label: str,
    left: dict[str, object],
    right: dict[str, object],
    *,
    left_name: str,
    right_name: str,
) -> None:
    for key in sorted(set(left) | set(right)):
        if left.get(key) != right.get(key):
            raise AssertionError(
                f"{label}: {key}: {left_name}={left.get(key)!r} != {right_name}={right.get(key)!r}"
            )


def _print_codegen_highlights(label: str, code: str) -> None:
    print(f"{label} — emitted defs of interest")
    interesting = [
        line
        for line in code.splitlines()
        if line.startswith("def _group_")
        or line.startswith("def cell_sheet1_")
        or line.startswith("    return _group_")
        or line.startswith("    '''Formula group")
    ]
    if not interesting:
        print("  (no group helpers / thin wrappers)")
        return
    for line in interesting:
        print(f"  {line}")
    print()


def _describe_groups(graph: DependencyGraph, report_keys: tuple[str, ...]) -> None:
    for key in report_keys:
        node = graph.get_node(key)
        assert node is not None
        print(f"  {key}")
        print(f"    kind / shape: {node.kind} / {node.shape}")
        print(f"    members:      {tuple(member_keys(node))}")
        print(f"    fingerprint:  {node.shape_fingerprint}")
        print(f"    skeleton:     {node.skeleton!r}")
        if node.skeleton is not None and node.member_bindings is not None:
            for member in member_keys(node):
                specialized = specialize_group(node.skeleton, node.member_bindings[member])
                print(f"    specialize({member}) -> {specialized!r}")


def main() -> None:
    print()
    print("Formula-group demo — row + column + individual cells")
    print("workbook → create_dependency_graph → coalesce → eval / codegen")
    print()

    workbook = write_demo_workbook(WORKBOOK)
    print(f"Wrote {workbook}")
    print(f"  row members:    {list(ROW_MEMBERS)}")
    print(f"  column members: {list(COL_MEMBERS)}")
    print(f"  cell targets:   {list(CELL_TARGETS)}")
    print()

    cell_graph = build_cell_only_graph(workbook)
    print("create_dependency_graph(..., formula_groups=False)")
    print(f"  nodes: {cell_graph.keys(order='workbook')}")
    print(f"  target_keys: {cell_graph.target_keys()}")
    print()

    group_graph = build_cell_only_graph(workbook)
    report = coalesce_formula_groups(group_graph)
    print("coalesce_formula_groups")
    print(f"  created_groups: {report.created_groups}")
    print(f"  skipped: {[(s.reason, s.members) for s in report.skipped_families]}")
    print(f"  nodes: {group_graph.keys(order='workbook')}")
    print(f"  target_keys (still member addresses): {group_graph.target_keys()}")
    print()

    assert len(report.created_groups) == 2, report.created_groups
    row_key, col_key = report.created_groups
    row_node = group_graph.get_node(row_key)
    col_node = group_graph.get_node(col_key)
    assert row_node is not None and col_node is not None
    # Contiguous row cover vs column cover (column RangeKeys use NodeKind.union).
    if row_node.shape is not NodeShape.row:
        row_key, col_key = col_key, row_key
        row_node, col_node = col_node, row_node
    assert row_node.kind is NodeKind.row and row_node.shape is NodeShape.row
    assert col_node.shape is NodeShape.column
    assert tuple(member_keys(row_node)) == ROW_MEMBERS
    assert tuple(member_keys(col_node)) == COL_MEMBERS
    for cell in CELL_TARGETS:
        node = group_graph.get_node(cell)
        assert node is not None and node.kind is NodeKind.cell

    print("Created groups")
    _describe_groups(group_graph, (row_key, col_key))
    print()
    print("Individual cells (unchanged occupancy)")
    for cell in CELL_TARGETS:
        loc = locate_cell(group_graph, cell)
        print(f"  {cell}: kind={group_graph.get_node(cell).kind} locate={loc.node_key if loc else None}")
    print()

    cell_eval = evaluate_targets(cell_graph, TARGETS)
    group_eval = evaluate_targets(group_graph, TARGETS)
    print("Evaluator (all targets)")
    for t in TARGETS:
        print(f"  {t}: cell-only={cell_eval[t]!r}  coalesced={group_eval[t]!r}")
    _assert_maps_equal(
        "evaluator cell-only vs coalesced",
        cell_eval,
        group_eval,
        left_name="cell-only",
        right_name="coalesced",
    )
    print("  ✓ evaluator matches with and without coalesce")
    print()

    cell_codegen, cell_code = codegen_targets(cell_graph, TARGETS)
    group_codegen, group_code = codegen_targets(group_graph, TARGETS)
    print("Codegen (all targets in one generate() call)")
    for t in TARGETS:
        print(f"  {t}: cell-only={cell_codegen[t]!r}  coalesced={group_codegen[t]!r}")
    _assert_maps_equal(
        "codegen cell-only vs coalesced",
        cell_codegen,
        group_codegen,
        left_name="cell-only",
        right_name="coalesced",
    )
    print("  ✓ codegen matches with and without coalesce")
    print()

    _assert_maps_equal(
        "cell-only eval vs codegen",
        cell_eval,
        cell_codegen,
        left_name="eval",
        right_name="codegen",
    )
    _assert_maps_equal(
        "coalesced eval vs codegen",
        group_eval,
        group_codegen,
        left_name="eval",
        right_name="codegen",
    )
    print("Parity")
    print("  ✓ cell-only: evaluator ↔ codegen")
    print("  ✓ coalesced: evaluator ↔ codegen")
    print()

    _print_codegen_highlights("Cell-only codegen", cell_code)
    _print_codegen_highlights("Coalesced codegen", group_code)
    group_helpers = [line for line in group_code.splitlines() if line.startswith("def _group_")]
    assert len(group_helpers) == 2, group_helpers
    assert "def _group_" not in cell_code

    print("Mermaid (coalesced — mixture of cells, row group, column group)")
    print(to_mermaid(group_graph))
    print()
    print("All checks passed.")


if __name__ == "__main__":
    main()
