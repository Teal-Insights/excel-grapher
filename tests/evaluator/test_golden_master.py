"""Simple golden master test without fixtures."""

from __future__ import annotations

import random
import time
from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import FormulaEvaluator, XlError, create_dependency_graph
from excel_grapher.evaluator.name_utils import normalize_address, parse_address
from tests.evaluator.discover_formula_cells import discover_formula_cells_in_rows
from tests.utils.modify_and_recalculate import (
    ExcelRecalculationError,
    modify_and_recalculate_workbook,
)

pytestmark = pytest.mark.slow

WORKBOOK_PATH = Path("example/data/lic-dsf-template-2025-08-12.xlsm")

# Sheet -> list of indicator row numbers (aligned with example workbook structure).
INDICATOR_CONFIG: dict[str, list[int]] = {
    "B1_GDP_ext": [35, 36, 39, 40],
    "B3_Exports_ext": [35, 36, 39, 40],
    "B4_other flows_ext": [35, 36, 39, 40],
}
MAX_DEPTH = 100
RTOL = 1e-5

# Number of leaf cells to perturb simultaneously
NUM_PERTURBATIONS = 50


def _format_cell_ref_for_excel(addr: str) -> str:
    """Convert internal cell address to Excel-style absolute reference."""
    if "!" not in addr:
        return addr

    sheet_name, cell_part = parse_address(addr)

    # Quote sheet if it has special characters (including parentheses for Excel)
    if " " in sheet_name or any(c in sheet_name for c in "-()"):
        quoted_sheet = f"'{sheet_name}'"
    else:
        quoted_sheet = sheet_name

    # Make absolute reference with $ signs
    cell_part = cell_part.replace("$", "")
    col = "".join(c for c in cell_part if c.isalpha())
    row = "".join(c for c in cell_part if c.isdigit())

    return f"{quoted_sheet}!${col}${row}"


def _read_numeric_cached_values(workbook_path: Path, addrs: list[str]) -> dict[str, float]:
    """Read cached numeric values for sheet-qualified addresses in one workbook load."""
    wb = fastpyxl.load_workbook(
        str(workbook_path),
        data_only=True,
        read_only=True,
        keep_vba=True,
    )
    try:
        out: dict[str, float] = {}
        for addr in addrs:
            if "!" not in addr:
                continue
            addr_n = normalize_address(addr)
            sheet_name, cell_part = parse_address(addr)
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            val = ws[cell_part.replace("$", "")].value
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                out[addr_n] = float(val)
        return out
    finally:
        wb.close()


def test_golden_master_inline(tmp_path: Path) -> None:
    """Inline golden master test."""
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")

    print("\n\nLoading graph...")
    start = time.time()
    targets: list[str] = []
    wb_f = fastpyxl.load_workbook(WORKBOOK_PATH, data_only=False, read_only=True, keep_vba=True)
    wb_v = fastpyxl.load_workbook(WORKBOOK_PATH, data_only=True, read_only=True, keep_vba=True)
    try:
        for sheet_name, rows in INDICATOR_CONFIG.items():
            targets.extend(
                discover_formula_cells_in_rows(
                    WORKBOOK_PATH, sheet_name, rows, wb_formulas=wb_f, wb_values=wb_v,
                )
            )
    finally:
        wb_f.close()
        wb_v.close()
    targets = [normalize_address(t) for t in targets]
    print(f"Discovered {len(targets)} targets in {time.time() - start:.1f}s")

    start = time.time()
    graph = create_dependency_graph(
        WORKBOOK_PATH,
        targets,
        load_values=True,
        max_depth=MAX_DEPTH,
    )
    print(f"Built graph with {len(graph)} nodes in {time.time() - start:.1f}s")

    # Compare only the configured indicator-row outputs (not the entire dependency closure).
    target_cells = list(dict.fromkeys(targets))
    print(f"Found {len(target_cells)} target cells in configured output rows")

    # Read original Excel values for comparison (to verify perturbation has effect)
    print("Reading original Excel values...")
    original_excel_values = _read_numeric_cached_values(WORKBOOK_PATH, target_cells)
    print(f"Read {len(original_excel_values)} original Excel values")

    # Find perturbable leaves
    leaves = [
        addr
        for addr in graph.leaf_keys()
        if (node := graph.get_node(addr))
        and isinstance(node.value, (int, float))
        and node.value != 0
    ]
    print(f"Found {len(leaves)} perturbable leaves")

    if not leaves:
        pytest.skip("No perturbable leaf cells found")

    # Pick multiple random leaves to perturb
    random.seed(42)
    num_to_perturb = min(NUM_PERTURBATIONS, len(leaves))
    selected_leaves = random.sample(leaves, num_to_perturb)

    # Build perturbations dict: {excel_ref: new_value}
    cell_modifications: dict[str, float] = {}
    perturbation_details: list[tuple[str, float, float]] = []  # (addr, old, new)

    for leaf_addr in selected_leaves:
        node = graph.get_node(leaf_addr)
        if node is None or not isinstance(node.value, (int, float)):
            continue
        # Random perturbation between -10% and +10%
        perturbation = random.uniform(-0.1, 0.1)
        new_value = node.value * (1 + perturbation)
        excel_ref = _format_cell_ref_for_excel(leaf_addr)
        cell_modifications[excel_ref] = new_value
        perturbation_details.append((leaf_addr, node.value, new_value))

    print(f"Perturbing {len(cell_modifications)} leaf cells (±10% each)")
    for addr, old, new in perturbation_details[:5]:  # Show first 5
        print(f"  {addr}: {old} -> {new}")
    if len(perturbation_details) > 5:
        print(f"  ... and {len(perturbation_details) - 5} more")

    # Modify Excel with all perturbations at once
    output_path = tmp_path / "modified.xlsm"

    try:
        print("Running LibreOffice recalculation...")
        start = time.time()
        modify_and_recalculate_workbook(
            input_path=WORKBOOK_PATH,
            output_path=output_path,
            cell_modifications=cell_modifications,
        )
        print(f"Recalculation completed in {time.time() - start:.1f}s")
    except (ExcelRecalculationError, RuntimeError) as e:
        pytest.skip(f"Excel recalculation not available: {e}")

    # Read perturbed Excel values
    print("Reading perturbed Excel values...")
    excel_values = _read_numeric_cached_values(output_path, target_cells)
    print(f"Read {len(excel_values)} perturbed Excel values")

    # Verify that perturbation actually changed at least one output
    changed_count = 0
    for addr in excel_values:
        if addr in original_excel_values:
            orig = original_excel_values[addr]
            new = excel_values[addr]
            if orig == 0:
                if new != 0:
                    changed_count += 1
            elif abs(new - orig) / abs(orig) > RTOL:
                changed_count += 1
    print(f"Perturbation changed {changed_count}/{len(excel_values)} output values")
    assert changed_count > 0, (
        f"Perturbation of {len(perturbation_details)} leaves did not change any output values - "
        "test is not validating anything meaningful"
    )

    # Compute Python values with all perturbations applied
    print("Computing Python values...")
    with FormulaEvaluator(graph) as ev:
        # Apply all perturbations
        for leaf_addr, _old, new_value in perturbation_details:
            ev.set_value(leaf_addr, new_value)
        results = ev.evaluate(list(excel_values.keys()))

    python_values = {
        addr: float(val)
        for addr, val in results.items()
        if isinstance(val, (int, float)) and not isinstance(val, XlError)
    }
    print(f"Computed {len(python_values)} Python values")

    # Compare
    matches = 0
    mismatches = 0
    for addr in excel_values:
        if addr not in python_values:
            mismatches += 1
            print(f"  Missing Python result: {addr} Excel={excel_values[addr]}")
            continue
        excel_val = excel_values[addr]
        python_val = python_values[addr]
        if excel_val == 0 and python_val == 0:
            matches += 1
        elif excel_val == 0:
            mismatches += 1
        else:
            rel_error = abs(python_val - excel_val) / abs(excel_val)
            if rel_error <= RTOL:
                matches += 1
            else:
                mismatches += 1
                print(f"  Mismatch: {addr} Excel={excel_val} Python={python_val}")

    print(f"\nResults: {matches} matches, {mismatches} mismatches")
    total = matches + mismatches
    accuracy = matches / total if total > 0 else 0
    print(f"Accuracy: {accuracy:.1%}")

    assert accuracy == 1.0, f"100% of golden master values must match with rtol {RTOL}. Measured accuracy: {accuracy:.1%}"
