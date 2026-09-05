"""Phase 2: formula bodies in internals.py call read_* for bound input leaves."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document


def _write_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    inputs = wb.add_worksheet("Inputs")
    calc = wb.add_worksheet("Calc")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        inputs.write(0, col, year)
        inputs.write_number(4, col, float(year))
    inputs.write("A5", "Primary balance")
    # Formula over a bound leaf (H5 = time_period 3).
    calc.write_formula("A1", "=Inputs!H5")
    # Unbound leaf reference.
    calc.write_formula("B1", "=Inputs!A1")
    # Binding-aligned full data_range.
    calc.write_formula("C1", "=SUM(Inputs!F5:J5)")
    # Partial / non-aligned slice of the same row.
    calc.write_formula("D1", "=SUM(Inputs!F5:H5)")
    wb.close()


BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.9.0",
    "workbook": "internals_readers.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Inputs",
            "data_range": "Inputs!F5:J5",
            "layout": "series",
            "input": {"setter": {"name": "set_borvelia_primary_balance"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                    }
                ],
            },
            "key": ["TIME_PERIOD"],
        }
    ],
}


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "internals_readers.xlsx"
    _write_workbook(path)
    return path


def test_single_file_generate_rewrites_bound_leaves(workbook: Path) -> None:
    """Single-file export shares the Phase 2 rewrite (not only modular internals)."""
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = ["Calc!A1", "Calc!B1", "Calc!C1", "Calc!D1"] + expand_data_range(
        "Inputs!F5:J5", workbook=workbook
    )
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )

    formula_section = code.split("# --- Formula cell functions ---", 1)[1].split(
        "def make_context", 1
    )[0]
    assert "read_borvelia_primary_balance(ctx, time_period=3)" in formula_section
    assert "xl_cell(ctx, 'Inputs!H5')" not in formula_section
    assert "read_borvelia_primary_balance_range(ctx)" in formula_section
    assert "xl_cell(ctx, 'Inputs!A1')" in formula_section
    assert "xl_range(ctx, 'Inputs!F5:H5')" in formula_section
    # Discovery metadata uses the same reader index as body rewrite.
    assert "read_borvelia_primary_balance(ctx, time_period=3)" in code
