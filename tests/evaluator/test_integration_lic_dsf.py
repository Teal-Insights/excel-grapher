"""
Integration test: evaluate LIC-DSF workbook formulas and compare with Excel cached values.
"""

import re
from pathlib import Path

import openpyxl
import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, XlError, create_dependency_graph
from tests.evaluator.discover_formula_cells import discover_formula_cells_in_rows

# This integration test loads a large workbook and can take a while to run.
pytestmark = pytest.mark.slow

# Path to the test workbook
WORKBOOK_PATH = Path("example/data/lic-dsf-template-2025-08-12.xlsm")

# Configuration matching the indicator mapping script
INDICATOR_CONFIG = {
    "B1_GDP_ext": [35, 36, 39, 40],
    "B3_Exports_ext": [35, 36, 39, 40],
    "B4_other flows_ext": [35, 36, 39, 40],
}


@pytest.fixture(scope="module")
def graph() -> DependencyGraph:
    """Load the LIC-DSF workbook and build DependencyGraph for indicator rows."""
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")

    # Discover formula cells in the configured rows
    targets: list[str] = []
    wb_f = openpyxl.load_workbook(WORKBOOK_PATH, data_only=False, read_only=True, keep_vba=True)
    wb_v = openpyxl.load_workbook(WORKBOOK_PATH, data_only=True, read_only=True, keep_vba=True)
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

    return create_dependency_graph(
        WORKBOOK_PATH,
        targets,
        load_values=True,
        max_depth=100,
    )


def extract_functions_from_formulas(graph: DependencyGraph) -> set[str]:
    """Extract all unique function names used in formulas."""
    functions: set[str] = set()
    func_pattern = re.compile(r"([A-Z][A-Z0-9_]*)\s*\(")

    for _key, node in graph.formula_nodes():
        if node.formula:
            for match in func_pattern.findall(node.formula):
                functions.add(match.upper())
        if node.normalized_formula:
            for match in func_pattern.findall(node.normalized_formula):
                functions.add(match.upper())

    return functions


def test_list_functions_used(graph: DependencyGraph) -> None:
    """Diagnostic test: list all functions used in the workbook."""
    functions = extract_functions_from_formulas(graph)
    print(f"\n\nFunctions used in LIC-DSF workbook ({len(functions)}):")
    for func in sorted(functions):
        print(f"  {func}")


def test_graph_loaded(graph: DependencyGraph) -> None:
    """Verify DependencyGraph was loaded successfully."""
    assert len(graph) > 0
    print(f"\n\nLoaded {len(graph)} cells from workbook")

    # Count formula cells vs leaf cells
    formula_cells = len(graph.formula_keys())
    leaf_cells = len(graph.leaf_keys())
    print(f"  Formula cells: {formula_cells}")
    print(f"  Leaf cells: {leaf_cells}")


def test_evaluate_formulas(graph: DependencyGraph) -> None:
    """Evaluate all formulas and compare with Excel cached values."""
    # Get list of formula cells that have cached numeric values (our targets)
    targets: list[str] = []
    for addr, node in graph.formula_nodes():
        if node.normalized_formula and isinstance(node.value, (int, float)):
            targets.append(addr)

    print(f"\n\nEvaluating {len(targets)} formula cells...")

    # Evaluate one at a time to catch errors
    matches = 0
    mismatches = 0
    errors = 0
    error_types: dict[str, int] = {}
    mismatch_details: list[tuple[str, float, float]] = []

    with FormulaEvaluator(graph) as ev:
        for addr in targets:
            node = graph.get_node(addr)
            assert node is not None
            expected = node.value
            try:
                computed = ev._evaluate_cell(addr)  # noqa: SLF001
            except NotImplementedError as e:
                errors += 1
                func_match = re.search(r"not implemented: (\w+)", str(e))
                if func_match:
                    func = func_match.group(1)
                    error_types[func] = error_types.get(func, 0) + 1
                continue
            except Exception as e:
                errors += 1
                err_type = type(e).__name__
                error_types[err_type] = error_types.get(err_type, 0) + 1
                continue

            if isinstance(computed, XlError):
                errors += 1
                key = f"XlError.{computed.name}"
                error_types[key] = error_types.get(key, 0) + 1
            elif computed is None:
                errors += 1
            elif isinstance(expected, (int, float)) and isinstance(
                computed, (int, float)
            ):
                # Compare with tolerance for floating point
                if abs(expected - computed) < 1e-6 or (
                    expected != 0 and abs((expected - computed) / expected) < 1e-6
                ):
                    matches += 1
                else:
                    mismatches += 1
                    if len(mismatch_details) < 10:  # Limit output
                        mismatch_details.append((addr, expected, computed))
            else:
                mismatches += 1

    print("\nResults:")
    print(f"  Matches: {matches}")
    print(f"  Mismatches: {mismatches}")
    print(f"  Errors: {errors}")

    if error_types:
        print("\n  Error types:")
        for err_type, count in sorted(error_types.items()):
            print(f"    {err_type}: {count}")

    if mismatch_details:
        print("\n  Sample mismatches (first 10):")
        for addr, expected, computed in mismatch_details:
            print(f"    {addr}: expected={expected}, computed={computed}")

    evaluated = matches + mismatches
    accuracy = matches / evaluated * 100 if evaluated else 0
    print(f"\n  Accuracy (of evaluated): {accuracy:.1f}% ({matches}/{evaluated})")
