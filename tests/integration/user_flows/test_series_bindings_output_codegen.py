"""Integration: CodeGenerator emits output compute functions from series bindings."""

from __future__ import annotations

import importlib
import sys
from collections.abc import Callable
from copy import deepcopy
from pathlib import Path
from typing import Any, cast

import xlsxwriter

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document

MICRO = Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks"
FIXTURES = Path(__file__).resolve().parents[3] / "tests" / "fixtures" / "series_bindings"


def _write_output_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year))
    ws.write_formula("G5", "=F5+1")
    wb.close()


BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.2.0",
    "workbook": "series_bindings_output.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "row_series",
            "input": {"setter": {"name": "set_borvelia_primary_balance"}},
            "output": {"compute": {"name": "compute_borvelia_primary_balance"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "concept": "REF_AREA",
                        "role": "key",
                        "scope": "series",
                        "bind": {"kind": "cell", "address": "Sheet1!A2", "read": "string"},
                        "include_in_record": False,
                    },
                    {
                        "concept": "INDICATOR",
                        "role": "key",
                        "scope": "series",
                        "bind": {
                            "kind": "row_label",
                            "label_column": "A",
                            "read": "string",
                            "normalize": "strip_trailing_unit",
                        },
                        "include_in_record": False,
                    },
                    {
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                    },
                ],
            },
            "key": ["TIME_PERIOD"],
            "series_context": {
                "REF_AREA": "Borvelia",
                "INDICATOR": "Primary balance (% of GDP)",
            },
        }
    ],
}


def test_codegen_includes_output_compute_and_setter(tmp_path: Path) -> None:
    workbook = tmp_path / "series_bindings_output.xlsx"
    _write_output_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = expand_data_range("Sheet1!F5:J5", workbook=workbook) + ["Sheet1!G5"]
    graph = create_dependency_graph(workbook, targets, load_values=True)

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )

    assert "def set_borvelia_primary_balance(" in code
    assert "def compute_borvelia_primary_balance(" in code
    assert "Record = dict[str, object]" in code
    assert "-> Records:" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = cast(Callable[[], Any], ns["make_context"])
    setter = cast(
        Callable[[Any, list[dict[str, object]]], None], ns["set_borvelia_primary_balance"]
    )
    compute = cast(Callable[..., list[dict[str, object]]], ns["compute_borvelia_primary_balance"])

    ctx = make_context()
    setter(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
    records = compute(ctx=ctx)
    by_period = {cast(int, r["TIME_PERIOD"]): r for r in records}
    assert by_period[4]["OBS_VALUE"] == 7.5
    assert by_period[5]["OBS_VALUE"] == 5.0


def test_generate_modules_exports_output_compute(tmp_path: Path) -> None:
    workbook = tmp_path / "series_bindings_output.xlsx"
    _write_output_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = expand_data_range("Sheet1!F5:J5", workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)

    files = CodeGenerator(graph).generate_modules(
        targets,
        package_name="exported_series_output",
        series_bindings=bindings,
        bindings_workbook=workbook,
    )
    assert "def compute_borvelia_primary_balance(" in files["exported_series_output/entrypoint.py"]
    assert "compute_borvelia_primary_balance" in files["exported_series_output/__init__.py"]
    assert "Record" in files["exported_series_output/entrypoint.py"]

    for relpath, content in files.items():
        out_path = tmp_path / relpath
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported_series_output")
        ctx = pkg.make_context()
        records = pkg.compute_borvelia_primary_balance(ctx=ctx)
        assert len(records) == 5
        assert all("OBS_VALUE" in r for r in records)
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "exported_series_output" or name.startswith("exported_series_output."):
                sys.modules.pop(name, None)
