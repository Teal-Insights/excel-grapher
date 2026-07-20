"""Integration: CodeGenerator emits output compute functions from series bindings."""

from __future__ import annotations

import importlib
import sys
from collections.abc import Callable
from copy import deepcopy
from pathlib import Path
from typing import Any, cast

import pytest
import xlsxwriter

from excel_grapher.exporter import (
    CodeGenerator,
    FieldDoc,
    SeriesFunctionDoc,
    register_series_docstring_callback,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document


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
    "schema_version": "1.3.0",
    "workbook": "series_bindings_output.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "series",
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


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "series_bindings_output.xlsx"
    _write_output_workbook(path)
    return path


def test_codegen_includes_output_compute_and_setter(workbook: Path) -> None:
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
    assert "def read_borvelia_primary_balance(" in code
    assert "def list_setters() -> list[str]:" in code
    assert "def list_readers() -> list[str]:" in code
    assert "def list_computes() -> list[str]:" in code
    assert "Record = dict[str, object]" in code
    assert "-> Records:" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = cast(Callable[[], Any], ns["make_context"])
    list_setters = cast(Callable[[], list[str]], ns["list_setters"])
    list_readers = cast(Callable[[], list[str]], ns["list_readers"])
    list_computes = cast(Callable[[], list[str]], ns["list_computes"])
    setter = cast(
        Callable[[Any, list[dict[str, object]]], None], ns["set_borvelia_primary_balance"]
    )
    reader = cast(Callable[..., object], ns["read_borvelia_primary_balance"])
    compute = cast(Callable[..., list[dict[str, object]]], ns["compute_borvelia_primary_balance"])

    assert list_setters() == ["set_borvelia_primary_balance"]
    assert list_readers() == ["read_borvelia_primary_balance"]
    assert list_computes() == ["compute_borvelia_primary_balance"]

    ctx = make_context()
    setter(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
    assert reader(ctx, time_period=4) == 7.5
    records = compute(ctx=ctx)
    by_period = {cast(int, r["TIME_PERIOD"]): r for r in records}
    assert by_period[4]["OBS_VALUE"] == 7.5
    assert by_period[5]["OBS_VALUE"] == 5.0


def test_generate_applies_series_docstring_callback_to_output_compute(
    workbook: Path,
) -> None:
    callback_name = "_test_integration_compute_docstring"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary=f"Compute {ctx.contract.series_id}.",
            purpose="Integration test purpose.",
            record_matching="One record per output cell.",
            field_descriptions={
                "TIME_PERIOD": FieldDoc(description="Reporting year."),
                "OBS_VALUE": FieldDoc(description="Computed value."),
            },
        ),
    )
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = expand_data_range("Sheet1!F5:J5", workbook=workbook) + ["Sheet1!G5"]
    graph = create_dependency_graph(workbook, targets, load_values=True)

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
            series_docstring_callback=callback_name,
        )

    ns: dict[str, object] = {}
    exec(code, ns)
    compute = cast(Callable[..., list[dict[str, object]]], ns["compute_borvelia_primary_balance"])
    assert compute.__doc__ is not None
    assert "Compute borvelia_primary_balance." in compute.__doc__
    assert "Examples:" in compute.__doc__


def test_generate_modules_exports_output_compute(
    tmp_path: Path,
    workbook: Path,
) -> None:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = expand_data_range("Sheet1!F5:J5", workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)

    files = CodeGenerator(graph).generate_modules(
        targets,
        series_bindings=bindings,
        bindings_workbook=workbook,
    )
    assert "def compute_borvelia_primary_balance(" in files["api.py"]
    assert "_readers.py" in files
    assert "def read_borvelia_primary_balance(" in files["_readers.py"]
    assert "def read_borvelia_primary_balance_range(" in files["_readers.py"]
    assert "from ._readers import" in files["api.py"]
    assert "def list_setters() -> list[str]:" in files["api.py"]
    assert "def list_readers() -> list[str]:" in files["api.py"]
    assert "def list_computes() -> list[str]:" in files["api.py"]
    assert "compute_borvelia_primary_balance" in files["__init__.py"]
    assert "read_borvelia_primary_balance" in files["__init__.py"]
    assert "read_borvelia_primary_balance_range" in files["__init__.py"]
    assert "list_setters" in files["__init__.py"]
    assert "list_readers" in files["__init__.py"]
    assert "list_computes" in files["__init__.py"]
    assert "Record" in files["api.py"] or "Record" in files.get("_api_helpers.py", "")

    pkg_dir = tmp_path / "exported_series_output"
    for filename, content in files.items():
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported_series_output")
        ctx = pkg.make_context()
        assert pkg.list_setters() == ["set_borvelia_primary_balance"]
        assert pkg.list_readers() == ["read_borvelia_primary_balance"]
        assert pkg.list_computes() == ["compute_borvelia_primary_balance"]
        records = pkg.compute_borvelia_primary_balance(ctx=ctx)
        assert len(records) == 5
        assert all("OBS_VALUE" in r for r in records)
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "exported_series_output" or name.startswith("exported_series_output."):
                sys.modules.pop(name, None)
