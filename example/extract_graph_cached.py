#!/usr/bin/env python3
"""
Map dependencies for LIC-DSF indicator rows using excel-grapher.

This script traces the dependency closure for key indicators across sheets
and validates against calcChain.xml.

Dynamic refs (OFFSET/INDIRECT) are resolved via a constraint-based config.
Iterative workflow: run the script; if DynamicRefError is raised, the message
includes the formula cell that needs a constraint. Inspect that cell and the
row/column headers in the workbook to decide plausible input domains, add the
address to LicDsfConstraints (with Annotated[int, Between(lo, hi)], Annotated[float, RealBetween(...)], or
Literal[...]), then re-run until the graph
builds.

The dependency graph is written to a pickle under ``.cache/`` (by default) when
inputs match; use ``--no-cache`` to force a full rebuild. If you change
``EXPORT_RANGES``/target logic without changing the workbook file, use
``--no-cache`` or bump ``GRAPH_CACHE_SCHEMA``.

With ``GRAPH_LOAD_VALUES=True``, each node’s ``value`` includes Excel’s cached
calculated result for formula cells (from a data-only workbook read at build
time), which ``main.py`` uses for an instant Figure 1 preview before
``FormulaEvaluator`` finishes.
"""

import argparse
import hashlib
import importlib.metadata
import pickle
import time
from pathlib import Path
from typing import (
    Literal,
    TypedDict,
)

import fastpyxl
import fastpyxl.utils.cell

from excel_grapher import (
    CycleError,
    DependencyGraph,
    DynamicRefError,
    create_dependency_graph,
    format_cell_key,
    get_calc_settings,
    to_graphviz,
    validate_graph,
)


class ExportRangeConfig(TypedDict):
    """
    Explicit range specification for export/annotation targets.

    Attributes:
        label: Human-readable label for the range (used for reporting only).
        range_spec: Sheet-qualified A1 range, e.g. "'Chart Data'!D10:D17".
        entrypoint_mode: Controls how export entrypoints are grouped for this
            range: "row_group" (one entrypoint per row) or "per_cell" (one
            entrypoint per cell, no row grouping).
    """

    label: str
    range_spec: str
    entrypoint_mode: Literal["row_group", "per_cell"]


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

WORKBOOK_PATH = Path("example/data/lic-dsf-template-2025-08-12.xlsm")
WORKBOOK_TEMPLATE_URL = "https://thedocs.worldbank.org/en/doc/f0ade6bcf85b6f98dbeb2c39a2b7770c-0360012025/original/LIC-DSF-IDA21-Template-08-12-2025-vf.xlsm"

GRAPH_CACHE_SCHEMA = 2
GRAPH_MAX_DEPTH = 50
GRAPH_LOAD_VALUES = True
GRAPH_USE_CACHED_DYNAMIC_REFS = True

_GRAPH_SCRIPT_DIR = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Export package
# ---------------------------------------------------------------------------

PACKAGE_NAME = "lic_dsf_2025_08_12"
EXPORT_DIR = Path("dist/lic-dsf-2025-08-12")

# ---------------------------------------------------------------------------
# Export ranges
# ---------------------------------------------------------------------------

STRESS_TEST_ROW_LABELS: list[str] = [
    "Baseline",
    "A1. Key variables at their historical averages in 2024-2034 2/",
    "B1. Real GDP growth",
    "B2. Primary balance",
    "B3. Exports",
    "B4. Other flows 3/",
    "B5. Depreciation",
    "B6. Combination of B1-B5",
    "",
    "C1. Combined contingent liabilities",
    "C2. Natural disaster",
    "C3. Commodity price",
    "C4. Market Financing",
    "A2. Alternative Scenario :[Customize, enter title]",
]


STRESS_TEST_BLOCKS: list[tuple[str, int]] = [
    ("PV of Debt-to-GDP Ratio", 239),
    ("PV of Debt-to-Revenue Ratio", 281),
    ("Debt Service-to-Revenue Ratio", 318),
    ("Debt Service-to-GDP Ratio", 351),
]


FIGURE_DATA_ROWS: list[int] = [
    # Figure 1 (Output 2-1 Stress_Charts_Ex)
    51,
    61,
    62,
    63,
    64,
    66,
    93,
    103,
    104,
    105,
    106,
    108,
    135,
    145,
    146,
    147,
    148,
    150,
    177,
    187,
    188,
    189,
    190,
    192,
    # Figure 2 extras (Output 2-2 Stress_Charts_Pub)
    263,
    264,
    265,
    267,
    306,
    341,
    342,
    343,
]


EXPORT_FIXED_RANGES: list[ExportRangeConfig] = [
    {
        "label": "External DSA risk rating signals",
        "range_spec": "'Chart Data'!D10:D17",
        "entrypoint_mode": "per_cell",
    },
    {
        "label": "Fiscal (Total Public Debt) risk rating signals",
        "range_spec": "'Chart Data'!I10:I14",
        "entrypoint_mode": "per_cell",
    },
    {
        "label": "Applicable tailored stress test signals",
        "range_spec": "'Chart Data'!I17:I19",
        "entrypoint_mode": "row_group",
    },
    {
        "label": "Fiscal space for moderate risk category",
        "range_spec": "'Chart Data'!E25:E27",
        "entrypoint_mode": "row_group",
    },
    {
        "label": "Overall rating",
        "range_spec": "'Chart Data'!L10:L11",
        "entrypoint_mode": "row_group",
    },
]


def _export_chart_data_ranges() -> list[ExportRangeConfig]:
    out: list[ExportRangeConfig] = list(EXPORT_FIXED_RANGES)
    seen_row_specs = {entry["range_spec"] for entry in out}

    def add_chart_data_row(row: int, label: str) -> None:
        range_spec = f"'Chart Data'!D{row}:X{row}"
        if range_spec in seen_row_specs:
            return
        out.append(
            {
                "label": label,
                "range_spec": range_spec,
                "entrypoint_mode": "row_group",
            }
        )
        seen_row_specs.add(range_spec)

    for metric_label, start_row in STRESS_TEST_BLOCKS:
        for i, row_label in enumerate(STRESS_TEST_ROW_LABELS):
            if not row_label:
                continue
            row = start_row + i
            add_chart_data_row(row, f"{metric_label} - {row_label}")

    for row in FIGURE_DATA_ROWS:
        add_chart_data_row(row, f"Figure data row {row}")

    return out


EXPORT_RANGES: list[ExportRangeConfig] = _export_chart_data_ranges()


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

    c1, r1 = fastpyxl.utils.cell.coordinate_from_string(start_a1)
    c2, r2 = fastpyxl.utils.cell.coordinate_from_string(end_a1)
    start_col_idx = fastpyxl.utils.cell.column_index_from_string(c1)
    end_col_idx = fastpyxl.utils.cell.column_index_from_string(c2)
    rlo, rhi = (r1, r2) if r1 <= r2 else (r2, r1)
    clo, chi = (
        (start_col_idx, end_col_idx)
        if start_col_idx <= end_col_idx
        else (end_col_idx, start_col_idx)
    )

    out: list[str] = []
    for row in range(rlo, rhi + 1):
        for col_idx in range(clo, chi + 1):
            col_letter = fastpyxl.utils.cell.get_column_letter(col_idx)
            out.append(format_cell_key(sheet, col_letter, row))
    return out


def _default_graph_cache_path(workbook: Path) -> Path:
    stem = workbook.resolve().stem
    return _GRAPH_SCRIPT_DIR / ".cache" / f"{stem}-dependency-graph.pkl"


def _targets_fingerprint(targets: list[str]) -> str:
    h = hashlib.sha256()
    for key in targets:
        h.update(key.encode("utf-8"))
        h.update(b"\n")
    return h.hexdigest()


def _excel_grapher_version() -> str:
    try:
        return importlib.metadata.version("excel-grapher")
    except importlib.metadata.PackageNotFoundError:
        return "unknown"


def _graph_cache_meta(
    workbook: Path,
    targets: list[str],
    *,
    max_depth: int,
) -> dict[str, object]:
    resolved = workbook.resolve()
    st = resolved.stat()
    return {
        "schema": GRAPH_CACHE_SCHEMA,
        "workbook": str(resolved),
        "workbook_mtime_ns": st.st_mtime_ns,
        "workbook_size": st.st_size,
        "targets_fingerprint": _targets_fingerprint(targets),
        "use_cached_dynamic_refs": GRAPH_USE_CACHED_DYNAMIC_REFS,
        "max_depth": max_depth,
        "load_values": GRAPH_LOAD_VALUES,
        "excel_grapher_version": _excel_grapher_version(),
    }


def _cache_meta_matches(
    stored: dict[str, object],
    workbook: Path,
    targets: list[str],
    *,
    max_depth: int,
) -> bool:
    if stored.get("schema") != GRAPH_CACHE_SCHEMA:
        return False
    resolved = workbook.resolve()
    try:
        st = resolved.stat()
    except OSError:
        return False
    if stored.get("workbook") != str(resolved):
        return False
    if stored.get("workbook_mtime_ns") != st.st_mtime_ns:
        return False
    if stored.get("workbook_size") != st.st_size:
        return False
    if stored.get("targets_fingerprint") != _targets_fingerprint(targets):
        return False
    if stored.get("use_cached_dynamic_refs") != GRAPH_USE_CACHED_DYNAMIC_REFS:
        return False
    if stored.get("max_depth") != max_depth:
        return False
    if stored.get("load_values") != GRAPH_LOAD_VALUES:
        return False
    return stored.get("excel_grapher_version") == _excel_grapher_version()


def try_load_graph_cache(
    cache_path: Path,
    workbook: Path,
    targets: list[str],
    *,
    max_depth: int,
) -> DependencyGraph | None:
    if not cache_path.is_file():
        return None
    try:
        with cache_path.open("rb") as f:
            payload = pickle.load(f)
    except (OSError, pickle.UnpicklingError, EOFError, AttributeError, TypeError):
        return None
    if not isinstance(payload, tuple) or len(payload) != 2:
        return None
    meta, graph = payload
    if not isinstance(meta, dict) or not isinstance(graph, DependencyGraph):
        return None
    if not _cache_meta_matches(meta, workbook, targets, max_depth=max_depth):
        return None
    return graph


def save_graph_cache(
    cache_path: Path,
    graph: DependencyGraph,
    workbook: Path,
    targets: list[str],
    *,
    max_depth: int,
) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    meta = _graph_cache_meta(workbook, targets, max_depth=max_depth)
    tmp = cache_path.with_suffix(cache_path.suffix + ".tmp")
    with tmp.open("wb") as f:
        pickle.dump((meta, graph), f, protocol=pickle.HIGHEST_PROTOCOL)
    tmp.replace(cache_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="LIC-DSF indicator dependency mapping.")
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="Ignore disk cache and rebuild the dependency graph.",
    )
    parser.add_argument(
        "--cache-path",
        type=Path,
        default=None,
        help="Pickle path for the graph cache (default: .cache/<workbook-stem>-dependency-graph.pkl).",
    )
    args = parser.parse_args()

    print("=" * 70)
    print("LIC-DSF Indicator Dependency Mapping")
    print("=" * 70)

    if not WORKBOOK_PATH.exists():
        print(f"Error: Workbook not found at {WORKBOOK_PATH}")
        return

    # Discover targets: explicit ranges (all cells) and indicator rows (formula cells only)
    print("\n1. Collecting target cells...")
    all_targets: list[str] = []

    for entry in EXPORT_RANGES:
        label = entry["label"]
        spec = entry["range_spec"]
        sheet_name, range_a1 = parse_range_spec(spec)
        targets = cells_in_range(sheet_name, range_a1)
        print(f"   {label}: {spec} -> {len(targets)} cells")
        all_targets.extend(targets)

    print(f"\n   Total targets: {len(all_targets)}")

    if not all_targets:
        print("No formula cells found. Exiting.")
        return

    cache_path = args.cache_path or _default_graph_cache_path(WORKBOOK_PATH)

    print("\n2. Loading / building dependency graph...", flush=True)
    t_graph = time.perf_counter()
    graph: DependencyGraph | None = None
    if not args.no_cache:
        graph = try_load_graph_cache(
            cache_path,
            WORKBOOK_PATH,
            all_targets,
            max_depth=GRAPH_MAX_DEPTH,
        )
        if graph is not None:
            print(f"   Loaded graph from cache: {cache_path}", flush=True)
            print(f"   Cache load time: {time.perf_counter() - t_graph:.2f}s", flush=True)

    if graph is None:
        print("   Starting create_dependency_graph...", flush=True)
        try:
            graph = create_dependency_graph(
                WORKBOOK_PATH,
                all_targets,
                load_values=GRAPH_LOAD_VALUES,
                max_depth=GRAPH_MAX_DEPTH,
                use_cached_dynamic_refs=GRAPH_USE_CACHED_DYNAMIC_REFS,
            )
        except DynamicRefError as e:
            print(f"\n   DynamicRefError: {e}")
            print(
                "   Add the reported cell's argument cells to LicDsfConstraints (address-style keys)"
                " using Annotated[..., Between(...)] or Annotated[..., RealBetween(...)] / Annotated[..., FromWorkbook()] as needed,"
                " then re-run. Or set USE_CACHED_DYNAMIC_REFS=True to resolve from cached values."
            )
            raise

        print(f"   Graph build time: {time.perf_counter() - t_graph:.2f}s", flush=True)
        if not args.no_cache:
            try:
                save_graph_cache(
                    cache_path,
                    graph,
                    WORKBOOK_PATH,
                    all_targets,
                    max_depth=GRAPH_MAX_DEPTH,
                )
                print(f"   Saved graph cache: {cache_path}", flush=True)
            except OSError as e:
                print(f"   Warning: could not write graph cache ({e})", flush=True)

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
    scope = {parse_range_spec(entry["range_spec"])[0] for entry in EXPORT_RANGES}
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
