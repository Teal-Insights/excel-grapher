"""Phase 2: formula bodies in internals.py call read_* for bound input leaves."""

from __future__ import annotations

import importlib
import sys
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


def _generate(workbook: Path) -> dict[str, str]:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = ["Calc!A1", "Calc!B1", "Calc!C1", "Calc!D1"] + expand_data_range(
        "Inputs!F5:J5", workbook=workbook
    )
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def test_internals_use_readers_for_bound_leaves_and_aligned_ranges(workbook: Path) -> None:
    files = _generate(workbook)
    internals = files["internals.py"]

    assert "read_borvelia_primary_balance(ctx, time_period=3)" in internals
    assert "xl_cell(ctx, 'Inputs!H5')" not in internals

    assert "read_borvelia_primary_balance_range(ctx)" in internals
    assert "xl_range(ctx, 'Inputs!F5:J5')" not in internals

    # Unbound leaf stays address-keyed.
    assert "xl_cell(ctx, 'Inputs!A1')" in internals
    # Partial slice is not binding-aligned.
    assert "xl_range(ctx, 'Inputs!F5:H5')" in internals

    assert "from ._readers import" in internals
    assert "read_borvelia_primary_balance" in internals


def test_internals_reader_calls_evaluate(
    workbook: Path,
    tmp_path: Path,
) -> None:
    files = _generate(workbook)
    pkg_dir = tmp_path / "internals_readers_pkg"
    pkg_dir.mkdir()
    for filename, content in files.items():
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("internals_readers_pkg")
        internals = importlib.import_module("internals_readers_pkg.internals")
        ctx = pkg.make_context()
        assert internals.cell_calc_a1(ctx) == 3.0
        assert internals.cell_calc_c1(ctx) == 15.0
        assert internals.cell_calc_d1(ctx) == 6.0
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "internals_readers_pkg" or name.startswith("internals_readers_pkg."):
                del sys.modules[name]


def test_groups_json_is_not_emitted(workbook: Path) -> None:
    """Sanity: unrelated bindings without groups still omit groups.json."""
    files = _generate(workbook)
    assert "groups.json" not in files
