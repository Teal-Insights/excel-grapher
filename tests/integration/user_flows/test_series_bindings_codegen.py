"""Integration: CodeGenerator emits Records setters from series bindings."""

from __future__ import annotations

import importlib
import sys
from collections.abc import Callable
from copy import deepcopy
from pathlib import Path
from typing import Any, cast

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document

MICRO = Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks"
WORKBOOK = MICRO / "series_bindings.xlsx"

BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.0.0",
    "workbook": "series_bindings.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "row_series",
            "editable": True,
            "setter": {"name": "set_borvelia_primary_balance"},
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


def test_codegen_includes_setters_and_updates_inputs() -> None:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=WORKBOOK))
    graph = create_dependency_graph(WORKBOOK, targets, load_values=True)

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=WORKBOOK,
        )

    assert "def set_borvelia_primary_balance(" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = cast(Callable[[], Any], ns["make_context"])
    set_borvelia_primary_balance = cast(
        Callable[[Any, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    ctx = make_context()
    set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
    assert ctx.inputs["Sheet1!I5"] == 7.5


def test_generate_modules_exports_series_binding_setters(tmp_path: Path) -> None:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=WORKBOOK))
    graph = create_dependency_graph(WORKBOOK, targets, load_values=True)

    files = CodeGenerator(graph).generate_modules(
        targets,
        series_bindings=bindings,
        bindings_workbook=WORKBOOK,
    )
    assert "def set_borvelia_primary_balance(" in files["api.py"]
    assert "set_borvelia_primary_balance" in files["__init__.py"]

    pkg_dir = tmp_path / "exported_series"
    for filename, content in files.items():
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported_series")
        ctx = pkg.make_context()
        pkg.set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
        assert ctx.inputs["Sheet1!I5"] == 7.5
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "exported_series" or name.startswith("exported_series."):
                sys.modules.pop(name, None)
