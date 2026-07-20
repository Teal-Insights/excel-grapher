"""Cell-only formula-family fixtures for Issue 3 coalesce / builder tests.

These graphs keep discrete formula cells (no multi-cell group yet). Passing them
through `coalesce_formula_groups` (or `create_dependency_graph(..., formula_groups=True)`)
should produce Issue 2–compatible unique-occupancy groups.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import xlsxwriter

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node


@dataclass(frozen=True)
class CellOnlyFamilyFixture:
    """Cell-only same-shape family ready to coalesce."""

    graph: DependencyGraph
    members: tuple[str, ...]
    leaf_keys: tuple[str, ...]


def build_cross_sheet_scale_cell_only() -> CellOnlyFamilyFixture:
    """Cross-sheet pair: `Sheet1!B1` / `Sheet2!B1` both `=…A1*10`."""
    members = ("Sheet1!B1", "Sheet2!B1")
    leaf_keys = ("Sheet1!A1", "Sheet2!A1")
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet2", "A", 1, value=2.0, is_leaf=True))
    for member, leaf, sheet, col, row, formula in (
        ("Sheet1!B1", "Sheet1!A1", "Sheet1", "B", 1, "=Sheet1!A1*10"),
        ("Sheet2!B1", "Sheet2!A1", "Sheet2", "B", 1, "=Sheet2!A1*10"),
    ):
        g.add_node(
            make_cell_node(
                sheet,
                col,
                row,
                formula=formula,
                normalized_formula=formula,
                is_leaf=False,
                is_target=True,
            )
        )
        g.add_edge(member, leaf)
    return CellOnlyFamilyFixture(graph=g, members=members, leaf_keys=leaf_keys)


def write_cross_sheet_scale_workbook(path: Path) -> Path:
    """Write a tiny workbook matching `build_cross_sheet_scale_cell_only` formulas."""
    path = Path(path)
    wb = xlsxwriter.Workbook(path)
    s1 = wb.add_worksheet("Sheet1")
    s2 = wb.add_worksheet("Sheet2")
    s1.write_number("A1", 1.0)
    s2.write_number("A1", 2.0)
    s1.write_formula("B1", "=Sheet1!A1*10", None, 10.0)
    s2.write_formula("B1", "=Sheet2!A1*10", None, 20.0)
    wb.close()
    return path
