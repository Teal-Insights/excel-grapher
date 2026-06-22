"""Unit tests for generated series-binding output compute functions."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import pytest
import xlsxwriter

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict, xl_cell
from excel_grapher.series_bindings import (
    Records,
    WorkbookSeriesBindings,
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
    validate_bindings_document,
)
from excel_grapher.series_bindings.compute_codegen import emit_compute_function, emit_computes_block
from excel_grapher.series_bindings.docstrings import (
    FieldDoc,
    SeriesFunctionDoc,
    register_series_docstring_callback,
)

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def _write_formula_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 5.0)
    ws.write_formula("C2", "=B2*2")
    wb.close()


def _exec_compute(
    lines: list[str],
    *,
    resolver: Callable[[str], Any],
) -> dict[str, object]:
    import warnings

    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
        "xl_cell": xl_cell,
        "warnings": warnings,
        "make_context": lambda inputs=None: EvalContext(
            inputs=coerce_inputs_dict(inputs or {}),
            resolver=resolver,
        ),
    }
    exec("\n".join(lines), namespace)
    return namespace


def test_emit_compute_returns_records_with_obs_value(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula.xlsx"
    _write_formula_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2"], load_values=True)

    series = {
        "id": "scaled_output",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_scaled_output"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
            "dimensions": [
                {
                    "concept": "LABEL",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "scaled"},
                }
            ],
        },
        "key": ["LABEL"],
    }
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    lines = [
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
        *emit_compute_function(series, resolved),
    ]

    formula_impl = graph.get_node("Sheet1!C2")
    assert formula_impl is not None

    def resolver(address: str):
        if address == "Sheet1!C2":
            return lambda ctx: 10.0
        return None

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[..., Records], ns["compute_scaled_output"])
    records = compute()
    assert len(records) == 1
    assert records[0]["LABEL"] == "scaled"
    assert records[0]["OBS_VALUE"] == 10.0


def test_emit_compute_borvelia_includes_frequency(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    from tests.unit.series_bindings.test_resolve import _write_borvelia_workbook

    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "shard_borvelia_output.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    lines = [
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
        *emit_compute_function(series, resolved),
    ]

    def resolver(address: str):
        return lambda ctx: xl_cell(ctx, address)

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[..., Records], ns["compute_borvelia_primary_balance"])
    records = compute()
    assert len(records) == 5
    assert all(r["FREQUENCY"] == "A" for r in records)
    assert {r["TIME_PERIOD"] for r in records} == {1, 2, 3, 4, 5}


def test_emit_computes_block_warns_when_output_has_no_graph_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula.xlsx"
    _write_formula_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!B2"], load_values=True)
    bindings = cast(
        WorkbookSeriesBindings,
        {
            "schema_version": "1.3.0",
            "series": [
                {
                    "id": "scaled_output",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!C2",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_scaled_output"}},
                    "structure": {
                        "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                        "dimensions": [
                            {
                                "concept": "LABEL",
                                "role": "key",
                                "scope": "series",
                                "bind": {"kind": "constant", "value": "scaled"},
                            }
                        ],
                    },
                    "key": ["LABEL"],
                }
            ],
        },
    )
    export = frozenset(CodeGenerator(graph)._generate_parts(["Sheet1!B2"])["all_cells"])
    with pytest.warns(UserWarning, match="No resolved output cells"):
        lines = emit_computes_block(graph, wb_path, bindings, export_addresses=export)

    assert "def compute_scaled_output(" not in "\n".join(lines)


def test_emit_computes_block_intersects_with_export_closure(tmp_path: Path) -> None:
    wb_path = tmp_path / "series_output.xlsx"
    from tests.integration.user_flows.test_series_bindings_output_codegen import (
        BINDINGS_DOCUMENT,
        _write_output_workbook,
    )

    _write_output_workbook(wb_path)
    bindings = cast(WorkbookSeriesBindings, BINDINGS_DOCUMENT)
    graph = create_dependency_graph(wb_path, ["Sheet1!G5"], load_values=True)
    export = frozenset(
        CodeGenerator(graph)._generate_parts(["Sheet1!G5"], dependency_targets=["Sheet1!G5"])[
            "all_cells"
        ]
    )

    with pytest.warns(UserWarning, match="codegen export closure"):
        lines = emit_computes_block(
            graph,
            wb_path,
            bindings,
            export_addresses=export,
        )

    f5 = graph.get_node("Sheet1!F5")
    g5 = graph.get_node("Sheet1!G5")
    assert f5 is not None and g5 is not None

    def resolver(address: str):
        if address == "Sheet1!G5":
            return lambda ctx: 42.0
        if address == "Sheet1!F5":
            return lambda ctx: float(f5.value or 0)
        return None

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[..., Records], ns["compute_borvelia_primary_balance"])
    records = compute()
    assert len(records) == 2
    periods = {cast(int, r["TIME_PERIOD"]) for r in records}
    assert periods == {1, 2}
    by_period = {cast(int, r["TIME_PERIOD"]): r for r in records}
    assert by_period[2]["OBS_VALUE"] == 42.0


def test_codegen_generate_output_compute_uses_export_closure(tmp_path: Path) -> None:
    wb_path = tmp_path / "series_output.xlsx"
    from copy import deepcopy

    from excel_grapher.series_bindings import validate_bindings_document
    from tests.integration.user_flows.test_series_bindings_output_codegen import (
        BINDINGS_DOCUMENT,
        _write_output_workbook,
    )

    _write_output_workbook(wb_path)
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    graph = create_dependency_graph(wb_path, ["Sheet1!G5"], load_values=True)

    with CodeGenerator(graph) as gen, pytest.warns(UserWarning) as captured_warnings:
        code = gen.generate(
            ["Sheet1!G5"],
            series_bindings=bindings,
            bindings_workbook=wb_path,
        )

    warning_messages = [str(warning.message) for warning in captured_warnings]
    assert (
        warning_messages.count(
            "Skipped 3 cell(s) in data_range not included in codegen export closure"
        )
        == 2
    )
    assert "Skipped 1 cell(s) in data_range not graph input leaf cells" in warning_messages
    assert (
        "Skipped 4 cell(s) in data_range not included in codegen export closure"
        not in warning_messages
    )

    ns: dict[str, object] = {}
    exec(code, ns)
    compute = cast(Callable[..., Records], ns["compute_borvelia_primary_balance"])
    records = compute()
    assert len(records) == 2


def test_emit_compute_structured_docstring_callback(tmp_path: Path) -> None:
    callback_name = "_test_compute_structured_docstring"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary="Compute scaled output.",
            purpose="Returns scaled output records.",
            record_matching="One record per output cell.",
            field_descriptions={
                "LABEL": FieldDoc(description="Series label."),
                "OBS_VALUE": FieldDoc(description="Computed value."),
            },
        ),
    )
    wb_path = tmp_path / "formula.xlsx"
    _write_formula_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2"], load_values=True)
    series = {
        "id": "scaled_output",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_scaled_output"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
            "dimensions": [
                {
                    "concept": "LABEL",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "scaled"},
                }
            ],
        },
        "key": ["LABEL"],
    }
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    lines = [
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
        *emit_compute_function(
            series,
            resolved,
            graph=graph,
            workbook=wb_path,
            bindings={
                "schema_version": "1.3.0",
                "workbook": "formula.xlsx",
                "series": [series],
                "concept_scheme": {},
            },
            series_docstring_callback=callback_name,
        ),
    ]

    def resolver(address: str):
        if address == "Sheet1!C2":
            return lambda ctx: 10.0
        return None

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[..., Records], ns["compute_scaled_output"])
    assert compute.__doc__ is not None
    assert "Compute scaled output." in compute.__doc__
    assert "Args:" in compute.__doc__
    assert "Returns:" in compute.__doc__
    assert "Examples:" in compute.__doc__


def test_emit_compute_google_docstring_renderer(tmp_path: Path) -> None:
    callback_name = "_test_compute_google_docstring"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary="Compute scaled output.",
            purpose="Returns scaled output records.",
            record_matching="One record per output cell.",
            field_descriptions={
                "LABEL": FieldDoc(description="Series label."),
                "OBS_VALUE": FieldDoc(description="Computed value."),
            },
        ),
    )
    wb_path = tmp_path / "formula.xlsx"
    _write_formula_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2"], load_values=True)
    series = {
        "id": "scaled_output",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_scaled_output"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
            "dimensions": [
                {
                    "concept": "LABEL",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "scaled"},
                }
            ],
        },
        "key": ["LABEL"],
    }
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    lines = [
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
        *emit_compute_function(
            series,
            resolved,
            graph=graph,
            workbook=wb_path,
            bindings={
                "schema_version": "1.3.0",
                "workbook": "formula.xlsx",
                "series": [series],
                "concept_scheme": {},
            },
            series_docstring_callback=callback_name,
            docstring_renderer="google",
        ),
    ]

    def resolver(address: str):
        if address == "Sheet1!C2":
            return lambda ctx: 10.0
        return None

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[..., Records], ns["compute_scaled_output"])
    assert compute.__doc__ is not None
    assert "Examples:" in compute.__doc__
    assert "Returns:" in compute.__doc__
    assert "\n    Returns:" not in compute.__doc__
    assert "Example:" not in compute.__doc__


def test_emit_compute_callback_requires_codegen_context(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula.xlsx"
    _write_formula_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2"], load_values=True)
    series = {
        "id": "scaled_output",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_scaled_output"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
            "dimensions": [
                {
                    "concept": "LABEL",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "scaled"},
                }
            ],
        },
        "key": ["LABEL"],
    }
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    with pytest.raises(ValueError, match="requires graph, workbook, and bindings"):
        _ = emit_compute_function(
            series,
            resolved,
            series_docstring_callback="series_notes",
        )


def _write_datetime_output_workbook(path: Path) -> None:
    from datetime import datetime

    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    date_format = wb.add_format({"num_format": "yyyy-mm-dd"})
    ws.write_datetime(0, 1, datetime(2024, 1, 1), date_format)
    ws.write_number(1, 1, 5.0)
    wb.close()


def test_emit_compute_datetime_literal_in_static_record(tmp_path: Path) -> None:
    from datetime import datetime

    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_output_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2"], load_values=True)
    series = {
        "id": "calendar_output",
        "sheet": "Inputs",
        "data_range": "Inputs!B2",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_calendar_output"}},
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
                    "bind": {"kind": "column_header", "header_row": 1, "read": "datetime"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    lines = emit_compute_function(series, resolved)
    code = "\n".join(lines)
    assert "import datetime" in code
    assert "datetime.datetime(2024, 1, 1, 0, 0)" in code

    def resolver(address: str):
        if address == "Inputs!B2":
            return lambda ctx: 7.0
        return None

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[[], Records], ns["compute_calendar_output"])
    records = compute()
    assert records[0]["TIME_PERIOD"] == datetime(2024, 1, 1)
    assert records[0]["OBS_VALUE"] == 7.0


def test_emit_compute_calendar_year_columns_round_trips(tmp_path: Path) -> None:
    from datetime import datetime

    wb_path = tmp_path / "calendar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    date_format = wb.add_format({"num_format": "yyyy-mm-dd"})
    periods = [datetime(2024, 1, 1), datetime(2024, 2, 1)]
    for col_index, period in enumerate(periods, start=1):
        ws.write_datetime(0, col_index, period, date_format)
        ws.write_number(1, col_index, float(col_index * 10))
    wb.close()

    graph = create_dependency_graph(wb_path, ["Inputs!B2", "Inputs!C2"], load_values=True)
    series = {
        "id": "calendar_output",
        "sheet": "Inputs",
        "data_range": "Inputs!B2:C2",
        "layout": "series",
        "output": {"compute": {"name": "compute_calendar_output"}},
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
                    "bind": {"kind": "column_header", "header_row": 1, "read": "datetime"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    lines = [
        "import datetime",
        "",
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
        *emit_compute_function(series, resolved, include_datetime_import=False),
    ]

    values = {"Inputs!B2": 10.0, "Inputs!C2": 20.0}

    def resolver(address: str):
        if address in values:
            return lambda ctx: values[address]
        return None

    ns = _exec_compute(lines, resolver=resolver)
    compute = cast(Callable[[], Records], ns["compute_calendar_output"])
    records = {record["TIME_PERIOD"]: record for record in compute()}
    assert records[datetime(2024, 1, 1)]["OBS_VALUE"] == 10.0
    assert records[datetime(2024, 2, 1)]["OBS_VALUE"] == 20.0


def test_emit_computes_block_includes_datetime_import_when_needed(tmp_path: Path) -> None:
    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_output_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2"], load_values=True)
    bindings: WorkbookSeriesBindings = {
        "schema_version": "1.4.0",
        "workbook": "calendar.xlsx",
        "series": [
            {
                "id": "calendar_output",
                "sheet": "Inputs",
                "data_range": "Inputs!B2",
                "layout": "scalar",
                "output": {"compute": {"name": "compute_calendar_output"}},
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
                            "bind": {
                                "kind": "column_header",
                                "header_row": 1,
                                "read": "datetime",
                            },
                        }
                    ],
                },
                "key": ["TIME_PERIOD"],
            }
        ],
    }
    code = "\n".join(emit_computes_block(graph, wb_path, bindings))
    assert "import datetime" in code
    assert code.count("import datetime") == 1
    assert "datetime.datetime(2024, 1, 1, 0, 0)" in code


def test_emit_compute_matrix_evaluates_formula_outputs(tmp_path: Path) -> None:
    from tests.fixtures.series_bindings.matrix_helpers import (
        macro_matrix_bindings_document,
        write_matrix_explicit_workbook,
    )

    wb_path = tmp_path / "matrix_compute.xlsx"
    write_matrix_explicit_workbook(wb_path, use_formulas=True)
    bindings = validate_bindings_document(
        macro_matrix_bindings_document(direction="output", workbook="matrix_compute.xlsx")
    )
    targets = expand_data_range("Inputs!B3:D5", workbook=wb_path)
    graph = create_dependency_graph(wb_path, targets, load_values=True)

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=wb_path,
        )

    assert "def compute_macro_matrix(" in code
    ns: dict[str, object] = {}
    exec(code, ns)
    compute = cast(Callable[..., Records], ns["compute_macro_matrix"])
    records = compute()
    assert len(records) == 9
    by_key = {
        (record["INDICATOR"], record["TIME_PERIOD"]): record["OBS_VALUE"] for record in records
    }
    assert by_key[("GDP growth", 2024)] == pytest.approx(2.4)
    assert by_key[("Inflation", 2025)] == pytest.approx(5.8)
    assert by_key[("Debt", 2026)] == pytest.approx(107.6)


def test_emit_compute_matrix_bindings_module_smoke(tmp_path: Path) -> None:
    from excel_grapher.series_bindings.smoke import smoke_test_bindings_module
    from excel_grapher.series_bindings.validate import validate_series_bindings
    from excel_grapher.series_bindings.workflow import generate_bindings_modules
    from tests.fixtures.series_bindings.matrix_helpers import (
        macro_matrix_bindings_document,
        write_matrix_explicit_workbook,
    )

    wb_path = tmp_path / "matrix_compute_smoke.xlsx"
    write_matrix_explicit_workbook(wb_path, use_formulas=True)
    bindings = validate_bindings_document(
        macro_matrix_bindings_document(direction="output", workbook="matrix_compute_smoke.xlsx")
    )
    targets = expand_data_range("Inputs!B3:D5", workbook=wb_path)
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True

    files = generate_bindings_modules(
        graph,
        targets=targets,
        bindings=bindings,
        workbook=wb_path,
    )
    smoke_test_bindings_module(
        files,
        bindings=bindings,
        graph=graph,
        workbook=wb_path,
        module_dir=tmp_path / "bindings_module",
        package_name="bindings_module",
    )
