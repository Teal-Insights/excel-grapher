"""Integration: CodeGenerator emits override setters for formula cells."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import xlsxwriter

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import xl_cell
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document


def _write_override_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws_in = wb.add_worksheet("Inputs")
    ws_in.write_number("A1", 10)
    ws_eng = wb.add_worksheet("Engine")
    ws_eng.write_formula("B1", "=Inputs!A1+1")
    ws_out = wb.add_worksheet("Output")
    ws_out.write_formula("C1", "=Engine!B1*2")
    wb.close()


BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.6.0",
    "workbook": "formula_override.xlsx",
    "series": [
        {
            "id": "engine_override",
            "sheet": "Engine",
            "data_range": "Engine!B1",
            "layout": "scalar",
            "input": {
                "mode": "override",
                "setter": {"name": "set_engine_b1"},
            },
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


def test_codegen_override_setter_updates_formula_cell_and_downstream(tmp_path: Path) -> None:
    workbook = tmp_path / "formula_override.xlsx"
    _write_override_workbook(workbook)
    bindings = validate_bindings_document(BINDINGS_DOCUMENT)
    targets = ["Output!C1"]
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
    graph = create_dependency_graph(workbook, targets, load_values=True)

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )

    assert "def set_engine_b1(" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = cast(Callable[[], Any], ns["make_context"])
    set_engine_b1 = cast(Callable[[Any, object], None], ns["set_engine_b1"])

    ctx = make_context()
    assert ctx.inputs["Inputs!A1"] == 10

    set_engine_b1(ctx, 99)
    assert ctx.inputs["Engine!B1"] == 99
    assert "Engine!B1" not in ctx.cache
    assert "Output!C1" not in ctx.cache
    assert xl_cell(ctx, "Output!C1") == 198
