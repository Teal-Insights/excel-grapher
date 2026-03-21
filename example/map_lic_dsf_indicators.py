#!/usr/bin/env python3
"""
Map dependencies for LIC-DSF indicator rows using excel-grapher.

This script traces the dependency closure for key indicators across sheets
and validates against calcChain.xml.

Dynamic refs (OFFSET/INDIRECT) are resolved via a constraint-based config.
Iterative workflow: run the script; if DynamicRefError is raised, the message
includes the formula cell that needs a constraint. Inspect that cell and the
row/column headers in the workbook to decide plausible input domains, add the
address to LicDsfConstraints (with Annotated[int, Between(lo, hi)] or
Literal[...]), then re-run until the graph
builds.
"""

from pathlib import Path
from typing import (  # noqa: F401 - Annotated/Literal used when adding constraints
    Annotated,
    Any,
    Literal,
    TypedDict,
    cast,
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
from excel_grapher.grapher.dynamic_refs import (
    FromWorkbook,  # noqa: F401 - used when adding constraints
)

# Row labels for the multi-row stress-test blocks (same row layout in each block).
# Blank string means that row is skipped when splitting by row.
STRESS_TEST_ROW_LABELS: list[str] = [
    "Baseline",
    "A1. Key variables at their historical averages in 2024-2034 2/",
    "B1. Real GDP growth",
    "B2. Primary balance",
    "B3. Exports",
    "B4. Other flows 3/",
    "B5. Depreciation",
    "B6. Combination of B1-B5",
    "",  # blank row
    "C1. Combined contingent liabilities",
    "C2. Natural disaster",
    "C3. Commodity price",
    "C4. Market Financing",
    "A2. Alternative Scenario :[Customize, enter title]",
]

# Explicit cell ranges to extract (sheet-qualified A1 range, e.g. "'Chart Data'!D10:D17").
# All cells in each range are included as graph targets.
# Multi-row stress-test blocks are split by row using STRESS_TEST_ROW_LABELS (blank row skipped).
_CHART_DATA_FIXED: list[tuple[str, str]] = [
    ("External DSA risk rating signals", "'Chart Data'!D10:D17"),
    ("Fiscal (Total Public Debt) risk rating signals", "'Chart Data'!I10:I14"),
    ("Applicable tailored stress test signals", "'Chart Data'!I17:I19"),
    ("Fiscal space for moderate risk category", "'Chart Data'!E25:E27"),
    ("Overall rating", "'Chart Data'!L10:L11"),
]

_STRESS_TEST_BLOCKS: list[tuple[str, int]] = [
    ("PV of Debt-to-GDP Ratio", 239),
    ("PV of Debt-to-Revenue Ratio", 281),
    ("Debt Service-to-Revenue Ratio", 318),
    ("Debt Service-to-GDP Ratio", 351),
]


def _chart_data_ranges() -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = list(_CHART_DATA_FIXED)
    for metric_label, start_row in _STRESS_TEST_BLOCKS:
        for i, row_label in enumerate(STRESS_TEST_ROW_LABELS):
            if not row_label:
                continue
            row = start_row + i
            out.append((f"{metric_label} - {row_label}", f"'Chart Data'!D{row}:X{row}"))
    return out


CHART_DATA_RANGES: list[tuple[str, str]] = _chart_data_ranges()

LiteralType = cast(Any, Literal)


def constrain_constant_range(
    constraints: type[Any],
    data: dict[str, Any],
    workbook_path: Path,
    *,
    sheet_name: str,
    range_a1: str,
) -> None:
    """Fill ``constraints`` and ``data`` with Literal types from a rectangular cell range."""
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    try:
        ws = wb[sheet_name]
        for row in ws[range_a1]:
            for cell in row:
                if cell.value is None:
                    continue
                key = f"{sheet_name}!{cell.coordinate}"
                val = cell.value
                constraints.__annotations__[key] = LiteralType[val]
                data[key] = val
    finally:
        wb.close()


# Dated template; adjust filename if using a different snapshot.
WORKBOOK_PATH = Path("example/data/lic-dsf-template-2026-01-31.xlsm")

# Set True to resolve OFFSET/INDIRECT from cached workbook values (no constraints).
# Set False to use constraint-based resolution; add address-style keys below as you hit DynamicRefError.
USE_CACHED_DYNAMIC_REFS = False

# Constraint types for cells that feed OFFSET/INDIRECT. Keys must be address-style (e.g. "Sheet1!B1").
# Add entries when the script raises DynamicRefError: the message lists leaf cells that need
# constraints. Add each to __annotations__ (with Annotated[int, Between(lo, hi)],
# Annotated[..., FromWorkbook()], or Literal[...]) then re-run. Repeat until the graph builds.
class LicDsfConstraints(TypedDict, total=False):
    pass

# PV_Base!B9xx = CONCAT("$", A9xx, "$", $A$<row>) → INDIRECT($B9xx). Row-index cells A917, A941, A965 (fixed).
# Treat these as constants derived from the current workbook values.
LicDsfConstraints.__annotations__["PV_Base!A917"] = Literal[64]
LicDsfConstraints.__annotations__["PV_Base!A941"] = Literal[90]
LicDsfConstraints.__annotations__["PV_Base!A965"] = Literal[115]
# A918:A938, A942:A962, A966:A986 each has a single cached letter D, E, …, X.
for _start, _end in [(918, 939), (942, 963), (966, 987)]:
    for _row in range(_start, _end):
        _letter = chr(ord("D") + _row - _start)
        LicDsfConstraints.__annotations__[f"PV_Base!A{_row}"] = LiteralType[_letter]

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


def _constrain_lookup_countries(constraints: type[Any]) -> None:
    """Constrain lookup!C4:C73 — LIC-DSF eligible country names.

    Generated by example/codegen_lookup_countries.py.
    """
    _countries: list[tuple[int, str]] = [
        (4, 'Afghanistan'),
        (5, 'Bangladesh'),
        (6, 'Benin'),
        (7, 'Bhutan'),
        (8, 'Burkina Faso'),
        (9, 'Burundi'),
        (10, 'Cambodia'),
        (11, 'Cameroon'),
        (12, 'Cabo Verde'),
        (13, 'Central African Republic'),
        (14, 'Chad'),
        (15, 'Comoros'),
        (16, 'Congo, DR'),
        (17, 'Congo, Republic of'),
        (18, "Cote d'Ivoire"),
        (19, 'Djibouti'),
        (20, 'Dominica'),
        (21, 'Eritrea'),
        (22, 'Ethiopia'),
        (23, 'Gambia, The'),
        (24, 'Ghana'),
        (25, 'Grenada'),
        (26, 'Guinea'),
        (27, 'Guinea-Bissau'),
        (28, 'Guyana'),
        (29, 'Haiti'),
        (30, 'Honduras'),
        (31, 'Kenya'),
        (32, 'Kiribati'),
        (33, 'Kyrgyz Republic'),
        (34, 'Lao PDR'),
        (35, 'Lesotho'),
        (36, 'Liberia'),
        (37, 'Madagascar'),
        (38, 'Malawi'),
        (39, 'Maldives'),
        (40, 'Mali'),
        (41, 'Marshall Islands'),
        (42, 'Mauritania'),
        (43, 'Micronesia'),
        (44, 'Moldova'),
        (45, 'Mozambique'),
        (46, 'Myanmar'),
        (47, 'Nepal'),
        (48, 'Nicaragua'),
        (49, 'Niger'),
        (50, 'Papua New Guinea'),
        (51, 'Rwanda'),
        (52, 'Samoa'),
        (53, 'Sao Tome & Principe'),
        (54, 'Senegal'),
        (55, 'Sierra Leone'),
        (56, 'Solomon Islands'),
        (57, 'Somalia'),
        (58, 'South Sudan'),
        (59, 'St. Lucia'),
        (60, 'St. Vincent & the Grenadines'),
        (61, 'Sudan'),
        (62, 'Tajikistan'),
        (63, 'Tanzania'),
        (64, 'Timor-Leste'),
        (65, 'Togo'),
        (66, 'Tonga'),
        (67, 'Tuvalu'),
        (68, 'Uganda'),
        (69, 'Uzbekistan'),
        (70, 'Vanuatu'),
        (71, 'Yemen, Republic of'),
        (72, 'Zambia'),
        (73, 'Zimbabwe'),
    ]
    for _row, _name in _countries:
        constraints.__annotations__[f"lookup!C{_row}"] = LiteralType[_name]


_constrain_lookup_countries(LicDsfConstraints)


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

    print(f"\n   Total targets: {len(all_targets)}")
    
    if not all_targets:
        print("No formula cells found. Exiting.")
        return
    
    # Build dependency graph (constraint-based or cached for OFFSET/INDIRECT)
    print("\n2. Building dependency graph...")
    dynamic_refs: DynamicRefConfig | None = None
    if not USE_CACHED_DYNAMIC_REFS:
        dynamic_refs = DynamicRefConfig.from_constraints_and_workbook(
            LicDsfConstraints,
            WORKBOOK_PATH,
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
            " using Annotated[..., Between(...)] / Annotated[..., FromWorkbook()] as needed,"
            " then re-run. Or set USE_CACHED_DYNAMIC_REFS=True to resolve from cached values."
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
    scope = {parse_range_spec(spec)[0] for _label, spec in CHART_DATA_RANGES}
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
