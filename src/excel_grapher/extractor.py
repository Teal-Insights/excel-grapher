"""
Utilities for discovering formula cells in Excel workbooks.

This module provides helpers for scanning Excel workbooks to find formula cells
in specific rows, which can then be used as targets for dependency graph building.

Migration Guide (from CellDict to DependencyGraph)
--------------------------------------------------
The CellDict and CellInfo abstractions have been removed. Use DependencyGraph
directly instead - it already provides O(1) lookup by cell address and preserves
full graph structure.

Old API -> New API:

    # OLD: Build a CellDict
    cells = build_cell_dict(path, {"Sheet1": [1, 2]})

    # NEW: Build a DependencyGraph directly
    from excel_grapher import create_dependency_graph, discover_formula_cells_in_rows
    targets = discover_formula_cells_in_rows(path, "Sheet1", [1, 2])
    graph = create_dependency_graph(path, targets, load_values=True)

    # OLD: Access cell data via CellInfo
    info = cells["Sheet1!A1"]
    formula = info.formula
    value = info.value

    # NEW: Access via Node directly
    node = graph.get_node("Sheet1!A1")
    formula = node.formula
    value = node.value

    # OLD: Filter formula vs value cells
    formula_cells = cells.formula_cells()
    value_cells = cells.value_cells()
    formula_keys = cells.formula_keys()
    value_keys = cells.value_keys()

    # NEW: Use DependencyGraph methods
    for key, node in graph.formula_nodes():
        ...
    for key, node in graph.leaf_node_items():
        ...
    formula_keys = graph.formula_keys()
    leaf_keys = graph.leaf_keys()

    # OLD: Check if cell has formula
    if cells["Sheet1!A1"].is_formula:
        ...

    # NEW: Check node.formula directly
    node = graph.get_node("Sheet1!A1")
    if node.formula is not None:
        ...

Benefits of using DependencyGraph directly:
- No memory duplication (CellInfo was copying data already on Node)
- Full graph structure preserved (edges, guards, dependents)
- Access to graph traversal methods (dependencies(), dependents(), evaluation_order())
- Access to cycle detection (cycle_report())
"""

from pathlib import Path

import openpyxl
import openpyxl.utils.cell

from .parser import format_cell_key

__all__ = [
    "discover_formula_cells_in_rows",
]


def discover_formula_cells_in_rows(
    wb_path: Path,
    sheet_name: str,
    rows: list[int],
) -> list[str]:
    """
    Scan specified rows and return sheet-qualified addresses for formula cells.

    Only includes cells that contain formulas (start with '=') and whose cached
    calculated value is numeric.

    Args:
        wb_path: Path to the Excel workbook
        sheet_name: Name of the sheet to scan
        rows: List of row numbers to scan

    Returns:
        List of sheet-qualified cell addresses (e.g., "'Sheet Name'!A1")
    """
    wb_formulas = openpyxl.load_workbook(wb_path, data_only=False, keep_vba=True)
    wb_values = openpyxl.load_workbook(wb_path, data_only=True, keep_vba=True)
    try:
        if sheet_name not in wb_formulas.sheetnames:
            return []

        ws_formulas = wb_formulas[sheet_name]
        ws_values = wb_values[sheet_name]
        targets: list[str] = []

        for row in rows:
            max_col = ws_formulas.max_column or 1
            for col_idx in range(1, max_col + 1):
                cell_formula = ws_formulas.cell(row=row, column=col_idx)
                if isinstance(cell_formula.value, str) and cell_formula.value.startswith("="):
                    cached_value = ws_values.cell(row=row, column=col_idx).value
                    if not isinstance(cached_value, (int, float)) or isinstance(cached_value, bool):
                        continue
                    col_letter = openpyxl.utils.cell.get_column_letter(col_idx)
                    targets.append(format_cell_key(sheet_name, col_letter, row))

        return targets
    finally:
        wb_formulas.close()
        wb_values.close()
