#!/usr/bin/env python3
"""
Map dependencies for LIC-DSF indicator rows using excel-grapher.

This script traces the dependency closure for key indicator rows across
B1, B3, and B4 sheets and validates against calcChain.xml.

Dynamic refs (OFFSET/INDIRECT) are resolved via a constraint-based config.
Iterative workflow: run the script; if DynamicRefError is raised, the message
includes the formula cell that needs a constraint. Inspect that cell and the
row/column headers in the workbook to decide plausible input domains, add the
address to LicDsfConstraints (with Annotated[int, Between(lo, hi)] or
Literal[...]) and to LIC_DSF_CONSTRAINTS_DATA, then re-run until the graph
builds.
"""

from pathlib import Path
from typing import (  # noqa: F401 - Annotated/Literal used when adding constraints
    Annotated,
    Literal,
    TypedDict,
)

import openpyxl
import openpyxl.utils.cell

from excel_grapher import (
    CycleError,
    DynamicRefConfig,
    DynamicRefError,
    create_dependency_graph,
    format_cell_key,
    get_calc_settings,
    to_graphviz,
    validate_graph,
)
from excel_grapher.core.cell_types import Between  # noqa: F401 - used when adding constraints


# Configuration: sheets and indicator rows to trace (row-based discovery)
class IndicatorConfig(TypedDict):
    sheet: str
    indicator_rows: list[int]


INDICATOR_CONFIG: list[IndicatorConfig] = [
    # Which, if any, rows do we want from "Baseline - external" and "A1_historical_ext"?
    # Six standardized stress tests:
    {"sheet": "B1_GDP_ext", "indicator_rows": [35, 36, 39, 40]},
    {"sheet": "B3_Exports_ext", "indicator_rows": [35, 36, 39, 40]},
    {"sheet": "B4_other flows_ext", "indicator_rows": [35, 36, 39, 40]},
    {"sheet": "B5_depreciation_ext", "indicator_rows": [35, 36, 39, 40]},
    {"sheet": "B6_Combo_mkt_ext", "indicator_rows": [35, 36, 39, 40]},
    {"sheet": "B6_Combo_non-mkt_ext", "indicator_rows": [35, 36, 39, 40]},
    # Mechanical external debt risk rating
    # Mechanical total public debt risk rating
]

# Explicit cell ranges to extract (sheet-qualified A1 range, e.g. "'Chart Data'!D10:D17").
# All cells in each range are included as graph targets.
CHART_DATA_RANGES: list[tuple[str, str]] = [
    ("External DSA risk rating signals", "'Chart Data'!D10:D17"),
    ("Fiscal (Total Public Debt) risk rating signals", "'Chart Data'!I10:I14"),
    ("Applicable tailored stress test signals", "'Chart Data'!I17:I19"),
    ("Fiscal space for moderate risk category", "'Chart Data'!E25:E27"),
    ("Overall rating", "'Chart Data'!L10:L11"),
    ("PV of Debt-to-GDP Ratio for all stress tests", "'Chart Data'!D239:X252"),
    ("PV of Debt-to-Revenue Ratio for all stress tests", "'Chart Data'!D281:X294"),
    ("Debt Service-to-Revenue Ratio for all stress tests", "'Chart Data'!D318:X331"),
    ("Debt Service-to-GDP Ratio for all stress tests", "'Chart Data'!D351:X364"),
]

# Dated template; adjust filename if using a different snapshot.
WORKBOOK_PATH = Path("example/data/lic-dsf-template-2026-01-31.xlsm")

# Set True to resolve OFFSET/INDIRECT from cached workbook values (no constraints).
# Set False to use constraint-based resolution; add address-style keys below as you hit DynamicRefError.
USE_CACHED_DYNAMIC_REFS = False

# Constraint types for cells that feed OFFSET/INDIRECT. Keys must be address-style (e.g. "Sheet1!B1").
# Add entries when the script raises DynamicRefError: the message lists leaf cells that need
# constraints. Add each to __annotations__ (with Annotated[int, Between(lo, hi)] or Literal[...])
# and to LIC_DSF_CONSTRAINTS_DATA, then re-run. Repeat until the graph builds.
class LicDsfConstraints(TypedDict, total=False):
    pass

# PV_Base!B9xx = CONCAT("$", A9xx, "$", $A$<row>) → INDIRECT($B9xx). Row-index cells A917, A941, A965 (fixed).
LicDsfConstraints.__annotations__["PV_Base!A917"] = Literal[64]
LicDsfConstraints.__annotations__["PV_Base!A941"] = Literal[90]
LicDsfConstraints.__annotations__["PV_Base!A965"] = Literal[115]
# A918:A938, A942:A962, A966:A986 each has a single cached letter D, E, …, X.
for _start, _end in [(918, 939), (942, 963), (966, 987)]:
    for _row in range(_start, _end):
        _letter = chr(ord("D") + _row - _start)
        LicDsfConstraints.__annotations__[f"PV_Base!A{_row}"] = Literal[_letter]

# Language selector and lookup table (feed INDIRECT/VLOOKUP for language-dependent refs).
# START!L10 = VLOOKUP(K10, lookup!BB4:BC7, 2); evaluator does not support VLOOKUP, so L10 is constrained too.
_LANG = Literal["English", "French", "Portuguese", "Spanish"]
_LANG_LOOKUP = Literal[
    "English", "French", "Portuguese", "Spanish", "Français", "Portugues", "Español"
]
LicDsfConstraints.__annotations__["START!L10"] = _LANG
LicDsfConstraints.__annotations__["START!K10"] = _LANG
for _r in range(4, 8):
    for _c in ("BB", "BC"):
        LicDsfConstraints.__annotations__[f"lookup!{_c}{_r}"] = _LANG_LOOKUP

LIC_DSF_CONSTRAINTS_DATA: dict[str, int | str | float] = {
    "PV_Base!A917": 64,
    "PV_Base!A941": 90,
    "PV_Base!A965": 115,
    **{f"PV_Base!A{r}": chr(ord("D") + r - _start)
      for _start, _end in [(918, 939), (942, 963), (966, 987)] for r in range(_start, _end)},
    "START!L10": "English",
    "START!K10": "English",
    **{f"lookup!{c}{r}": "English" for r in range(4, 8) for c in ("BB", "BC")},
}


def parse_range_spec(spec: str) -> tuple[str, str]:
    """
    Parse a sheet-qualified range spec into (sheet_name, range_a1).

    Accepts specs like "'Chart Data'!D10:D17" or "Sheet1!A1:B2".
    """
    if "!" not in spec:
        raise ValueError(f"Range spec must contain '!': {spec!r}")
    sheet_part, range_part = spec.split("!", 1)
    sheet_part = sheet_part.strip()
    if sheet_part.startswith("'") and sheet_part.endswith("'"):
        sheet_part = sheet_part[1:-1].replace("''", "'")
    return sheet_part, range_part.strip()


def cells_in_range(sheet: str, range_a1: str) -> list[str]:
    """
    Expand an A1 range to a list of sheet-qualified cell keys.

    range_a1 may be a single cell ("D10") or a range ("D10:D17", "D239:X252").
    """
    if ":" in range_a1:
        start_a1, end_a1 = range_a1.split(":", 1)
        start_a1 = start_a1.strip()
        end_a1 = end_a1.strip()
    else:
        start_a1 = end_a1 = range_a1.strip()

    c1, r1 = openpyxl.utils.cell.coordinate_from_string(start_a1)
    c2, r2 = openpyxl.utils.cell.coordinate_from_string(end_a1)
    start_col_idx = openpyxl.utils.cell.column_index_from_string(c1)
    end_col_idx = openpyxl.utils.cell.column_index_from_string(c2)
    rlo, rhi = (r1, r2) if r1 <= r2 else (r2, r1)
    clo, chi = (start_col_idx, end_col_idx) if start_col_idx <= end_col_idx else (end_col_idx, start_col_idx)

    out: list[str] = []
    for row in range(rlo, rhi + 1):
        for col_idx in range(clo, chi + 1):
            col_letter = openpyxl.utils.cell.get_column_letter(col_idx)
            out.append(format_cell_key(sheet, col_letter, row))
    return out


def discover_formula_cells_in_rows(
    wb_path: Path,
    sheet_name: str,
    rows: list[int],
) -> list[str]:
    """
    Scan specified rows and return sheet-qualified addresses for formula cells.

    Includes every cell that contains a formula (value starts with '='). Uses
    excel_grapher's format_cell_key so keys match the dependency graph.
    Cached values are not used so workbooks with no default inputs (formulas
    returning errors) are still discovered.
    """
    wb_formulas = openpyxl.load_workbook(wb_path, data_only=False, keep_vba=True)
    try:
        if sheet_name not in wb_formulas.sheetnames:
            print(f"  Warning: Sheet '{sheet_name}' not found")
            return []

        ws_formulas = wb_formulas[sheet_name]
        targets: list[str] = []

        for row in rows:
            max_col = ws_formulas.max_column or 1
            for col_idx in range(1, max_col + 1):
                cell_formula = ws_formulas.cell(row=row, column=col_idx)
                if isinstance(cell_formula.value, str) and cell_formula.value.startswith("="):
                    col_letter = openpyxl.utils.cell.get_column_letter(col_idx)
                    targets.append(format_cell_key(sheet_name, col_letter, row))

        return targets
    finally:
        wb_formulas.close()


def main() -> None:
    print("=" * 70)
    print("LIC-DSF Indicator Dependency Mapping")
    print("=" * 70)
    
    if not WORKBOOK_PATH.exists():
        print(f"Error: Workbook not found at {WORKBOOK_PATH}")
        return
    
    # Discover targets: explicit ranges (all cells) and indicator rows (formula cells only)
    print("\n1. Collecting target cells...")
    all_targets: list[str] = []

    for label, spec in CHART_DATA_RANGES:
        sheet_name, range_a1 = parse_range_spec(spec)
        targets = cells_in_range(sheet_name, range_a1)
        print(f"   {label}: {spec} -> {len(targets)} cells")
        all_targets.extend(targets)

    print("   Indicator rows (formula cells):")
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
    
    # Build dependency graph (constraint-based or cached for OFFSET/INDIRECT)
    print("\n2. Building dependency graph...")
    dynamic_refs: DynamicRefConfig | None = None
    if not USE_CACHED_DYNAMIC_REFS:
        dynamic_refs = DynamicRefConfig.from_constraints(
            LicDsfConstraints, LIC_DSF_CONSTRAINTS_DATA
        )
    try:
        graph = create_dependency_graph(
            WORKBOOK_PATH,
            all_targets,
            load_values=False,
            max_depth=50,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=USE_CACHED_DYNAMIC_REFS,
        )
    except DynamicRefError as e:
        print(f"\n   DynamicRefError: {e}")
        print(
            "   Add the reported cell's argument cells to LicDsfConstraints (address-style keys)"
            " and LIC_DSF_CONSTRAINTS_DATA, then re-run. Or set USE_CACHED_DYNAMIC_REFS=True to"
            " resolve from cached values."
        )
        raise

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
        print(
            f"   Example must-cycle path: {' -> '.join(report.example_must_cycle_path)}"
        )
    if report.example_may_cycle_path:
        print(
            f"   Example may-cycle path:  {' -> '.join(report.example_may_cycle_path)}"
        )
    
    # Validate against calcChain.xml
    print("\n5. Validating against calcChain.xml...")
    scope = {config["sheet"] for config in INDICATOR_CONFIG} | {
        parse_range_spec(spec)[0] for _label, spec in CHART_DATA_RANGES
    }
    result = validate_graph(graph, WORKBOOK_PATH, scope=scope)
    
    print(f"   Valid: {result.is_valid}")
    for msg in result.messages:
        print(f"   {msg}")
    
    if result.in_graph_not_in_chain:
        print(
            f"\n   Cells in graph but not in calcChain ({len(result.in_graph_not_in_chain)}):"
        )
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
