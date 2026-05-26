"""Integration: CodeGenerator emits Records setters from series bindings."""

from __future__ import annotations

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
    set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 4, "value": 7.5}])
    assert ctx.inputs["Sheet1!I5"] == 7.5
