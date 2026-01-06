#!/usr/bin/env python3
"""
Map dependencies for LIC-DSF indicator rows using excel-grapher.

This script traces the dependency closure for key indicator rows across
B1, B3, and B4 sheets and validates against calcChain.xml.
"""

from pathlib import Path
from typing import TypedDict

import openpyxl
import openpyxl.utils.cell

from excel_grapher import (
    CycleError,
    create_dependency_graph,
    get_calc_settings,
    to_graphviz,
    validate_graph,
)


# Configuration: sheets and indicator rows to trace
class IndicatorConfig(TypedDict):
    sheet: str
    indicator_rows: list[int]


INDICATOR_CONFIG: list[IndicatorConfig] = [
    {"sheet": "B1_GDP_ext", "indicator_rows": [35, 36, 39, 40]},
    {"sheet": "B3_Exports_ext", "indicator_rows": [35, 36, 39, 40]},
    {"sheet": "B4_other flows_ext", "indicator_rows": [35, 36, 39, 40]},
]

WORKBOOK_PATH = Path("example/data/Gold-Standard-LIC-DSF.xlsm")


def discover_formula_cells_in_rows(
    wb_path: Path,
    sheet_name: str,
    rows: list[int],
) -> list[str]:
    """
    Scan specified rows and return sheet-qualified addresses for formula cells.

    Only includes cells that contain formulas (start with '=') and whose cached
    calculated value is numeric.
    """
    wb_formulas = openpyxl.load_workbook(wb_path, data_only=False, keep_vba=True)
    wb_values = openpyxl.load_workbook(wb_path, data_only=True, keep_vba=True)
    try:
        if sheet_name not in wb_formulas.sheetnames or sheet_name not in wb_values.sheetnames:
            print(f"  Warning: Sheet '{sheet_name}' not found")
            return []

        ws_formulas = wb_formulas[sheet_name]
        ws_values = wb_values[sheet_name]
        targets: list[str] = []

        for row in rows:
            # Scan all columns up to max_column
            max_col = ws_formulas.max_column or 1
            for col_idx in range(1, max_col + 1):
                cell_formula = ws_formulas.cell(row=row, column=col_idx)
                if isinstance(cell_formula.value, str) and cell_formula.value.startswith("="):
                    cached_value = ws_values.cell(row=row, column=col_idx).value
                    if not isinstance(cached_value, (int, float)) or isinstance(cached_value, bool):
                        continue
                    col_letter = openpyxl.utils.cell.get_column_letter(col_idx)
                    targets.append(f"{sheet_name}!{col_letter}{row}")

        return targets
    finally:
        wb_formulas.close()
        wb_values.close()


def main() -> None:
    print("=" * 70)
    print("LIC-DSF Indicator Dependency Mapping")
    print("=" * 70)
    
    if not WORKBOOK_PATH.exists():
        print(f"Error: Workbook not found at {WORKBOOK_PATH}")
        print("Make sure Gold-Standard-LIC-DSF.xlsm is in the project root.")
        return
    
    # Discover all formula cells in indicator rows
    print("\n1. Discovering formula cells in indicator rows...")
    all_targets: list[str] = []
    
    for config in INDICATOR_CONFIG:
        sheet = config["sheet"]
        rows = config["indicator_rows"]
        print(f"   {sheet}: rows {rows}")
        
        targets = discover_formula_cells_in_rows(WORKBOOK_PATH, sheet, rows)
        print(f"      Found {len(targets)} formula cells")
        all_targets.extend(targets)
    
    print(f"\n   Total targets: {len(all_targets)}")
    
    if not all_targets:
        print("No formula cells found. Exiting.")
        return
    
    # Build dependency graph
    print("\n2. Building dependency graph...")
    graph = create_dependency_graph(
        WORKBOOK_PATH,
        all_targets,
        load_values=False,  # Skip cached values for speed
        max_depth=50,
    )
    
    print(f"   Nodes in graph: {len(graph)}")
    print(f"   Leaf nodes: {sum(1 for _ in graph.leaves())}")
    print(f"   Formula nodes: {len(graph) - sum(1 for _ in graph.leaves())}")
    
    # Group nodes by sheet
    sheets: dict[str, int] = {}
    for key in graph:
        node = graph.get_node(key)
        if node:
            sheets[node.sheet] = sheets.get(node.sheet, 0) + 1
    
    print("\n   Nodes by sheet:")
    for sheet_name in sorted(sheets.keys()):
        print(f"      {sheet_name}: {sheets[sheet_name]}")

    # Workbook calc settings (useful context for interpreting cycles)
    print("\n3. Workbook calculation settings...")
    settings = get_calc_settings(WORKBOOK_PATH)
    print(f"   Iterate enabled: {settings.iterate_enabled}")
    print(f"   Iterate count:   {settings.iterate_count}")
    print(f"   Iterate delta:   {settings.iterate_delta}")

    # Cycle analysis (must-cycle vs may-cycle)
    print("\n4. Cycle analysis...")
    report = graph.cycle_report()
    print(f"   Must-cycles: {len(report.must_cycles)}")
    print(f"   May-cycles:  {len(report.may_cycles)}")
    if report.example_must_cycle_path:
        print(f"   Example must-cycle path: {' -> '.join(report.example_must_cycle_path)}")
    if report.example_may_cycle_path:
        print(f"   Example may-cycle path:  {' -> '.join(report.example_may_cycle_path)}")
    
    # Validate against calcChain.xml
    print("\n5. Validating against calcChain.xml...")
    scope = {config["sheet"] for config in INDICATOR_CONFIG}
    result = validate_graph(graph, WORKBOOK_PATH, scope=scope)
    
    print(f"   Valid: {result.is_valid}")
    for msg in result.messages:
        print(f"   {msg}")
    
    if result.in_graph_not_in_chain:
        print(f"\n   Cells in graph but not in calcChain ({len(result.in_graph_not_in_chain)}):")
        for cell in sorted(result.in_graph_not_in_chain)[:10]:
            print(f"      {cell}")
        if len(result.in_graph_not_in_chain) > 10:
            print(f"      ... and {len(result.in_graph_not_in_chain) - 10} more")
    
    # Evaluation order stats
    print("\n6. Computing evaluation order...")
    try:
        # Non-strict mode will warn and exclude nodes involved in may-cycles, but
        # still fails on must-cycles.
        order = graph.evaluation_order(strict=False)
        print(f"   Evaluation order computed: {len(order)} nodes")
        print(f"   First 5 (leaves): {order[:5]}")
        print(f"   Last 5 (targets): {order[-5:]}")
    except CycleError as e:
        kind = "must-cycle" if e.is_must_cycle else "may-cycle"
        print(f"   Error ({kind}): {e}")
        if e.cycle_path:
            print(f"   Cycle path: {' -> '.join(e.cycle_path)}")
    
    # Optional: save a small subgraph visualization
    print("\n7. Sample visualization (first target's immediate deps)...")
    if all_targets:
        sample_target = all_targets[0]
        sample_deps = graph.dependencies(sample_target)
        print(f"   {sample_target} depends on {len(sample_deps)} cells:")
        for dep in sorted(sample_deps)[:5]:
            guard = graph.edge_attrs(sample_target, dep).get("guard")
            if guard is None:
                print(f"      {dep}")
            else:
                print(f"      {dep}  [guarded: {guard}]")
        if len(sample_deps) > 5:
            print(f"      ... and {len(sample_deps) - 5} more")

        # Emit a DOT snippet for quick inspection (guarded edges render dashed + labeled).
        try:
            dot = to_graphviz(graph, highlight={sample_target}, rankdir="LR")
            print("\n   GraphViz DOT (truncated to first ~40 lines):")
            for line in dot.splitlines()[:40]:
                print(f"      {line}")
            if len(dot.splitlines()) > 40:
                print("      ...")
        except Exception as e:
            print(f"   Could not render GraphViz DOT: {e}")
    
    print("\n" + "=" * 70)
    print("Done.")


if __name__ == "__main__":
    main()
