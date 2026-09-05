"""Integration: CodeGenerator emits Records setters from series bindings."""

from __future__ import annotations

from collections.abc import Callable
from copy import deepcopy
from pathlib import Path
from typing import Any, cast

import pytest

from excel_grapher.exporter import (
    CodeGenerator,
    FieldDoc,
    SeriesFunctionDoc,
    register_series_docstring_callback,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document
from tests.integration.user_flows.utils import write_series_bindings_workbook

BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.3.0",
    "workbook": "series_bindings.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "series",
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


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "series_bindings.xlsx"
    write_series_bindings_workbook(path)
    return path


def test_codegen_includes_setters_and_updates_inputs(workbook: Path) -> None:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
    graph = create_dependency_graph(workbook, targets, load_values=True)

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )

    assert "def set_borvelia_primary_balance(" in code
    assert "def read_borvelia_primary_balance(" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = cast(Callable[[], Any], ns["make_context"])
    list_setters = cast(Callable[[], list[str]], ns["list_setters"])
    list_readers = cast(Callable[[], list[str]], ns["list_readers"])
    list_computes = cast(Callable[[], list[str]], ns["list_computes"])
    set_borvelia_primary_balance = cast(
        Callable[[Any, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    read_borvelia_primary_balance = cast(
        Callable[..., object],
        ns["read_borvelia_primary_balance"],
    )

    assert list_setters() == ["set_borvelia_primary_balance"]
    assert list_readers() == ["read_borvelia_primary_balance"]
    assert list_computes() == []

    ctx = make_context()
    set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
    assert ctx.inputs["Sheet1!I5"] == 7.5
    assert read_borvelia_primary_balance(ctx, time_period=4) == 7.5


def test_generate_applies_series_docstring_callback(workbook: Path) -> None:
    callback_name = "_test_integration_setter_docstring"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary=f"Set {ctx.contract.series_id}.",
            purpose="Integration test purpose.",
            record_matching="Match by TIME_PERIOD.",
            field_descriptions={
                "TIME_PERIOD": FieldDoc(description="Reporting year."),
                "OBS_VALUE": FieldDoc(description="Observed value."),
            },
        ),
    )
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
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
    setter = cast(
        Callable[[Any, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    assert setter.__doc__ is not None
    assert "Set borvelia_primary_balance." in setter.__doc__
    assert "Args:" in setter.__doc__
    assert "Returns:" in setter.__doc__
    assert "Required record fields:" in setter.__doc__


def test_generate_accepts_direct_series_docstring_callback(workbook: Path) -> None:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
    graph = create_dependency_graph(workbook, targets, load_values=True)

    def callback(ctx: Any) -> SeriesFunctionDoc:
        return SeriesFunctionDoc(
            summary=f"Set {ctx.contract.series_id} directly.",
            purpose="Integration test direct callback purpose.",
            record_matching="Match by TIME_PERIOD.",
            field_descriptions={
                "TIME_PERIOD": FieldDoc(description="Reporting year."),
                "OBS_VALUE": FieldDoc(description="Observed value."),
            },
        )

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
            series_docstring_callback=callback,
        )

    ns: dict[str, object] = {}
    exec(code, ns)
    setter = cast(
        Callable[[Any, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    assert setter.__doc__ is not None
    assert "Set borvelia_primary_balance directly." in setter.__doc__


def test_generate_applies_google_docstring_renderer(workbook: Path) -> None:
    callback_name = "_test_integration_setter_google_docstring"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary=f"Set {ctx.contract.series_id}.",
            purpose="Integration test purpose.",
            record_matching="Match by TIME_PERIOD.",
            field_descriptions={
                "TIME_PERIOD": FieldDoc(description="Reporting year."),
                "OBS_VALUE": FieldDoc(description="Observed value."),
            },
        ),
    )
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
    graph = create_dependency_graph(workbook, targets, load_values=True)

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
            series_docstring_callback=callback_name,
            docstring_renderer="google",
        )

    ns: dict[str, object] = {}
    exec(code, ns)
    setter = cast(
        Callable[[Any, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    assert setter.__doc__ is not None
    assert "Args:" in setter.__doc__
    assert "Returns:" in setter.__doc__
    assert "Required record fields:" in setter.__doc__
