"""Phase 2: formula bodies rewrite bound constant leaves to read_*."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document


def _write_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    engine = wb.add_worksheet("Engine")
    engine.write_number("C5", 2021)
    engine.write_formula("C10", "=IF(C5>=2020,1,0)")
    wb.close()


BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.11.0",
    "workbook": "constant_readers.xlsx",
    "series": [
        {
            "id": "shock_year_anchor",
            "sheet": "Engine",
            "data_range": "Engine!C5",
            "layout": "scalar",
            "constant": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [],
            },
            "key": [],
        }
    ],
}


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "constant_readers.xlsx"
    _write_workbook(path)
    return path


def test_single_file_rewrites_constant_leaves(workbook: Path) -> None:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = ["Engine!C10", "Engine!C5"]
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
    assert "read_shock_year_anchor(ctx)" in formula_section
    assert "xl_cell(ctx, 'Engine!C5')" not in formula_section
    assert "def set_shock_year_anchor" not in code
    assert "def read_shock_year_anchor" in code
