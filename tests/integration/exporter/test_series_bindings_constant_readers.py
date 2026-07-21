"""Phase 2: formula bodies rewrite bound constant leaves to read_*."""

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


def _generate(workbook: Path) -> dict[str, str]:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = ["Engine!C10", "Engine!C5"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def test_internals_rewrite_constant_leaves_to_readers(workbook: Path) -> None:
    files = _generate(workbook)
    assert "_readers.py" in files
    internals = files["internals.py"]
    assert "read_shock_year_anchor(ctx)" in internals
    assert "xl_cell(ctx, 'Engine!C5')" not in internals
    assert "from ._readers import" in internals

    api = files["api.py"]
    assert "def list_setters() -> list[str]:" in api
    assert "return []" in api.split("def list_setters")[1].split("def ")[0]
    assert "read_shock_year_anchor" in api
    assert "def set_shock_year_anchor" not in api


def test_constant_reader_evaluates(workbook: Path, tmp_path: Path) -> None:
    files = _generate(workbook)
    pkg_dir = tmp_path / "constant_readers_pkg"
    pkg_dir.mkdir()
    for filename, content in files.items():
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("constant_readers_pkg")
        internals = importlib.import_module("constant_readers_pkg.internals")
        ctx = pkg.make_context()
        assert internals.cell_engine_c10(ctx) == 1.0
        assert pkg.read_shock_year_anchor(ctx) == 2021
        assert pkg.list_setters() == []
        assert "read_shock_year_anchor" in pkg.list_readers()
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "constant_readers_pkg" or name.startswith("constant_readers_pkg."):
                del sys.modules[name]


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
