"""Unit tests for generated series-binding setters."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import cast

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)
from excel_grapher.series_bindings.docstrings import (
    FieldDoc,
    SeriesFunctionDoc,
    register_series_docstring_callback,
)
from excel_grapher.series_bindings.setter_codegen import (
    emit_input_coerce_helpers,
    emit_setter_function,
    emit_setter_helpers,
    emit_setters_block,
)
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def _write_borvelia_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year - 3))
    wb.close()


def _exec_setters(
    lines: list[str],
    *,
    extra: dict[str, object] | None = None,
) -> dict[str, object]:
    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
    }
    if extra:
        namespace.update(extra)
    source_lines = lines
    if "def coerce_setter_input(" not in "\n".join(lines):
        source_lines = emit_input_coerce_helpers() + emit_setter_helpers() + lines
    exec("\n".join(source_lines), namespace)
    return namespace


def test_emit_setter_updates_context_by_key(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"TIME_PERIOD": 3, "OBS_VALUE": 42.0}])
    assert ctx.inputs["Inputs!H5"] == 42.0


def test_emit_setter_missing_key_raises(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    ns = _exec_setters(emit_setter_function(series, resolved))
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    with pytest.raises(ValueError, match="missing key fields"):
        setter(ctx, [{"OBS_VALUE": 1.0}])


def test_emit_setter_missing_measure_raises(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    ns = _exec_setters(emit_setter_function(series, resolved))
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    with pytest.raises(ValueError, match="missing required field 'OBS_VALUE'"):
        setter(ctx, [{"TIME_PERIOD": 3}])


def test_emit_setter_allowed_fields_literal_is_alphabetically_sorted(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    allowed_line = next(line for line in lines if line.strip().startswith("allowed_fields="))
    assert (
        allowed_line == "        allowed_fields=frozenset("
        "{'INDICATOR', 'OBS_VALUE', 'REF_AREA', 'TIME_PERIOD', 'UNIT_MEASURE'}),"
    )


def test_emit_setter_allowed_fields_literal_includes_address_fields_in_order(
    tmp_path: Path,
) -> None:
    wb_path = tmp_path / "dup.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("C1", 1)
    ws.write_number("D1", 1)
    ws.write_number("C2", 10)
    ws.write_number("D2", 20)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Sheet1!C2", "Sheet1!D2"], load_values=True)
    series = {
        "id": "dup_headers",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2:D2",
        "layout": "series",
        "setter": {"name": "set_dup_headers", "allow_address": True, "strict": False},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
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
        "validation": {"require_unique_key": True},
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    allowed_line = next(line for line in lines if line.strip().startswith("allowed_fields="))
    assert allowed_line == (
        "        allowed_fields=frozenset({'OBS_VALUE', 'TIME_PERIOD', 'address', 'cell_address'}),"
    )


def test_emit_setter_allow_address(tmp_path: Path) -> None:
    wb_path = tmp_path / "dup.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("C1", 1)
    ws.write_number("D1", 1)
    ws.write_number("C2", 10)
    ws.write_number("D2", 20)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Sheet1!C2", "Sheet1!D2"], load_values=True)
    series = {
        "id": "dup_headers",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2:D2",
        "layout": "series",
        "setter": {"name": "set_dup_headers", "allow_address": True, "strict": False},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
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
        "validation": {"require_unique_key": True},
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    assert resolved["requires_address"] is True
    lines = emit_setter_function(series, resolved)
    ns = _exec_setters(lines)
    setter = cast(Callable[[EvalContext, list[dict[str, object]]], None], ns["set_dup_headers"])
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"address": "Sheet1!D2", "OBS_VALUE": 99.0}])
    assert ctx.inputs["Sheet1!D2"] == 99.0


def test_emit_setters_block_emits_helpers_once(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    code = "\n".join(emit_setters_block(graph, wb_path, bindings))

    assert code.count("def _apply_series_records(") == 1
    assert code.count("def _coerce_records(") == 1
    assert code.count("def coerce_setter_input(") == 1
    assert code.count("if TYPE_CHECKING:") == 1
    assert (
        code.count(
            "SeriesInput = Records | Record | Sequence[Scalar] | pd.DataFrame | pl.DataFrame"
        )
        >= 1
    )
    assert "_KEY_ORDER_BORVELIA_PRIMARY_BALANCE" in code
    assert "_apply_series_records(" in code
    assert "def set_borvelia_primary_balance(" in code


def test_emit_setters_block_all_series(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    lines = emit_setters_block(graph, wb_path, bindings)
    assert "def set_borvelia_primary_balance(" in "\n".join(lines)


def test_emit_setters_block_skips_series_without_graph_leaf_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!A2"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    with pytest.warns(UserWarning, match="No resolved input cells"):
        lines = emit_setters_block(graph, wb_path, bindings)

    assert "def set_borvelia_primary_balance(" not in "\n".join(lines)


def test_emit_setter_structured_docstring_callback(tmp_path: Path) -> None:
    callback_name = "_test_setter_structured_docstring"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary="Set borvelia values.",
            purpose="Updates borvelia primary balance inputs.",
            record_matching="Match records by TIME_PERIOD.",
            field_descriptions={
                "TIME_PERIOD": FieldDoc(description="Reporting year."),
                "OBS_VALUE": FieldDoc(description='Value with "quotes".'),
            },
        ),
    )
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(
        series,
        resolved,
        graph=graph,
        workbook=wb_path,
        bindings=bindings,
        series_docstring_callback=callback_name,
    )
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    assert setter.__doc__ is not None
    assert "Set borvelia values." in setter.__doc__
    assert 'Value with "quotes".' in setter.__doc__
    exec(
        "\n".join(emit_input_coerce_helpers() + emit_setter_helpers() + lines),
        {"EvalContext": EvalContext, "coerce_inputs_dict": coerce_inputs_dict},
    )


def test_emit_setter_google_docstring_renderer(tmp_path: Path) -> None:
    callback_name = "_test_setter_google_docstring"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary="Set borvelia values.",
            purpose="Updates borvelia primary balance inputs.",
            record_matching="Match records by TIME_PERIOD.",
            field_descriptions={
                "TIME_PERIOD": FieldDoc(description="Reporting year."),
                "OBS_VALUE": FieldDoc(description="Observed value."),
            },
        ),
    )
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(
        series,
        resolved,
        graph=graph,
        workbook=wb_path,
        bindings=bindings,
        series_docstring_callback=callback_name,
        docstring_renderer="google",
    )
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    assert setter.__doc__ is not None
    assert "Args:" in setter.__doc__
    assert "Returns:" in setter.__doc__
    assert "Examples:" in setter.__doc__
    assert "Required record fields:" in setter.__doc__


def test_emit_setter_docstring_callback_none_omits_docstring(tmp_path: Path) -> None:
    callback_name = "_test_setter_none_docstring"
    register_series_docstring_callback(callback_name, lambda ctx: None)
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(
        series,
        resolved,
        graph=graph,
        workbook=wb_path,
        bindings=bindings,
        series_docstring_callback=callback_name,
    )
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_borvelia_primary_balance"],
    )
    assert setter.__doc__ is None


def test_emit_setter_callback_requires_codegen_context(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    with pytest.raises(ValueError, match="requires graph, workbook, and bindings"):
        _ = emit_setter_function(
            series,
            resolved,
            series_docstring_callback="series_notes",
        )


def _write_bool_key_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Flags")
    ws.write_boolean("A1", True)
    ws.write_boolean("B1", False)
    ws.write_number("A2", 10.0)
    ws.write_number("B2", 20.0)
    wb.close()


def test_emit_setter_bool_key_round_trips(tmp_path: Path) -> None:
    wb_path = tmp_path / "flags.xlsx"
    _write_bool_key_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Flags!A2", "Flags!B2"], load_values=True)
    series = {
        "id": "bool_keyed",
        "sheet": "Flags",
        "data_range": "Flags!A2:B2",
        "layout": "series",
        "setter": {"name": "set_bool_keyed"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "concept": "SLOT",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "bool"},
                }
            ],
        },
        "key": ["SLOT"],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    code = "\n".join(lines)
    assert "(('SLOT', True),): 'Flags!A2'" in code
    assert "(('SLOT', False),): 'Flags!B2'" in code

    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_bool_keyed"],
    )
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"SLOT": True, "OBS_VALUE": 99.0}])
    setter(ctx, [{"SLOT": False, "OBS_VALUE": 88.0}])
    assert ctx.inputs["Flags!A2"] == 99.0
    assert ctx.inputs["Flags!B2"] == 88.0


def _write_datetime_key_workbook(path: Path) -> None:
    from datetime import datetime

    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    date_format = wb.add_format({"num_format": "yyyy-mm-dd"})
    periods = [datetime(2024, 1, 1), datetime(2024, 2, 1)]
    for col_index, period in enumerate(periods, start=1):
        ws.write_datetime(0, col_index, period, date_format)
        ws.write_number(1, col_index, float(col_index))
    wb.close()


def test_emit_setter_datetime_key_round_trips(tmp_path: Path) -> None:
    from datetime import datetime

    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_key_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2", "Inputs!C2"], load_values=True)
    series = {
        "id": "calendar_keyed",
        "sheet": "Inputs",
        "data_range": "Inputs!B2:C2",
        "layout": "series",
        "setter": {"name": "set_calendar_keyed"},
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
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    code = "\n".join(lines)
    assert "datetime(2024, 1, 1, 0, 0)" in code
    assert "datetime(2024, 2, 1, 0, 0)" in code

    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_calendar_keyed"],
    )
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"TIME_PERIOD": datetime(2024, 1, 1), "OBS_VALUE": 11.0}])
    setter(ctx, [{"TIME_PERIOD": datetime(2024, 2, 1), "OBS_VALUE": 22.0}])
    assert ctx.inputs["Inputs!B2"] == 11.0
    assert ctx.inputs["Inputs!C2"] == 22.0


def test_emit_setters_block_includes_datetime_aliases_when_needed(tmp_path: Path) -> None:

    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_key_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2"], load_values=True)
    bindings: WorkbookSeriesBindings = {
        "schema_version": "1.4.0",
        "workbook": "calendar.xlsx",
        "series": [
            {
                "id": "calendar_keyed",
                "sheet": "Inputs",
                "data_range": "Inputs!B2",
                "layout": "scalar",
                "input": {"setter": {"name": "set_calendar_keyed"}},
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
    code = "\n".join(emit_setters_block(graph, wb_path, bindings))
    assert "from datetime import date, datetime, timedelta" in code
    assert "Scalar = str | int | float | bool | datetime | None" in code


def test_emit_setter_scalar_bool_measure_round_trips(tmp_path: Path) -> None:
    wb_path = tmp_path / "bool_scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Flags")
    ws.write_boolean("B2", True)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Flags!B2"], load_values=True)
    series = {
        "id": "bool_scalar_measure",
        "sheet": "Flags",
        "data_range": "Flags!B2",
        "layout": "scalar",
        "setter": {"name": "set_bool_scalar_measure"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "bool",
                "bind": {"kind": "data_cell", "read": "bool"},
            },
            "dimensions": [],
        },
        "key": [],
    }
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_bool_scalar_measure"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"OBS_VALUE": False}])
    assert ctx.inputs["Flags!B2"] is False


def test_emit_setter_scalar_bool_key_dimension_round_trips(tmp_path: Path) -> None:
    wb_path = tmp_path / "bool_scalar.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Flags")
    ws.write_boolean("B2", True)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Flags!B2"], load_values=True)
    concept_scheme = {
        "id": "flags",
        "concepts": [{"id": "IS_ACTIVE", "dtype": "bool"}],
    }
    series = {
        "id": "bool_scalar_key",
        "sheet": "Flags",
        "data_range": "Flags!B2",
        "layout": "scalar",
        "setter": {"name": "set_bool_scalar_key"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "bool",
                "bind": {"kind": "data_cell", "read": "bool"},
            },
            "dimensions": [
                {
                    "concept": "IS_ACTIVE",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": True},
                }
            ],
        },
        "key": ["IS_ACTIVE"],
    }
    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=concept_scheme,
    )
    lines = emit_setter_function(series, resolved)
    code = "\n".join(lines)
    assert "(('IS_ACTIVE', True),): 'Flags!B2'" in code

    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_bool_scalar_key"],
    )
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"IS_ACTIVE": True, "OBS_VALUE": False}])
    assert ctx.inputs["Flags!B2"] is False


def test_emit_setters_block_calendar_year_setter_round_trips(tmp_path: Path) -> None:
    from datetime import datetime

    wb_path = tmp_path / "calendar.xlsx"
    _write_datetime_key_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!B2", "Inputs!C2"], load_values=True)
    bindings: WorkbookSeriesBindings = {
        "schema_version": "1.4.0",
        "workbook": "calendar.xlsx",
        "concept_scheme": {
            "id": "calendar",
            "concepts": [{"id": "TIME_PERIOD", "dtype": "datetime"}],
        },
        "series": [
            {
                "id": "calendar_keyed",
                "sheet": "Inputs",
                "data_range": "Inputs!B2:C2",
                "layout": "series",
                "input": {"setter": {"name": "set_calendar_keyed"}},
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
    lines = emit_setters_block(graph, wb_path, bindings)
    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_calendar_keyed"],
    )
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [{"TIME_PERIOD": datetime(2024, 1, 1), "OBS_VALUE": 11.0}])
    setter(ctx, [{"TIME_PERIOD": datetime(2024, 2, 1), "OBS_VALUE": 22.0}])
    assert ctx.inputs["Inputs!B2"] == 11.0
    assert ctx.inputs["Inputs!C2"] == 22.0


def test_emit_setter_series_signature_uses_series_input(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    code = "\n".join(emit_setter_function(series, resolved))
    assert "records: SeriesInput," in code


def test_emit_setter_positional_values_update_context(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    ns = _exec_setters(emit_setter_function(series, resolved))
    setter = cast(Callable[[EvalContext, object], None], ns["set_borvelia_primary_balance"])

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, [-2.0, -1.0, 0.0, 7.5, 8.0])
    assert ctx.inputs["Inputs!I5"] == 7.5
    assert ctx.inputs["Inputs!J5"] == 8.0


def test_emit_setter_pandas_dataframe_updates_context(tmp_path: Path) -> None:
    pd = pytest.importorskip("pandas")
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    ns = _exec_setters(emit_setter_function(series, resolved))
    setter = cast(Callable[[EvalContext, object], None], ns["set_borvelia_primary_balance"])

    df = pd.DataFrame({"TIME_PERIOD": [4, 5], "OBS_VALUE": [7.5, 8.0]})
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, df)
    assert ctx.inputs["Inputs!I5"] == 7.5
    assert ctx.inputs["Inputs!J5"] == 8.0


def test_emit_setters_block_matrix_explicit_codegen_shape_and_round_trip(tmp_path: Path) -> None:
    from tests.fixtures.series_bindings.matrix_helpers import (
        MATRIX_EXPLICIT_BINDINGS,
        matrix_leaf_address_map,
        write_matrix_explicit_workbook,
    )

    wb_path = tmp_path / "matrix_inputs.xlsx"
    write_matrix_explicit_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!B3:D5"),
        load_values=True,
    )
    bindings = load_series_bindings(MATRIX_EXPLICIT_BINDINGS)
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    by_key = matrix_leaf_address_map(resolved)
    assert by_key[(("INDICATOR", "GDP growth"), ("TIME_PERIOD", 2024))] == "Inputs!B3"
    assert by_key[(("INDICATOR", "Debt"), ("TIME_PERIOD", 2026))] == "Inputs!D5"

    code = "\n".join(emit_setters_block(graph, wb_path, bindings))
    assert "def set_macro_matrix(" in code

    ns = _exec_setters(emit_setters_block(graph, wb_path, bindings))
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_macro_matrix"],
    )
    ctx = EvalContext(
        inputs=coerce_inputs_dict({"Inputs!B3": 1.2}),
        resolver=lambda _a: None,
    )
    setter(
        ctx,
        [
            {"INDICATOR": "GDP growth", "TIME_PERIOD": 2025, "OBS_VALUE": 9.9},
            {"INDICATOR": "Debt", "TIME_PERIOD": 2026, "OBS_VALUE": 44.4},
        ],
    )
    assert ctx.inputs["Inputs!C3"] == 9.9
    assert ctx.inputs["Inputs!D5"] == 44.4
    assert ctx.inputs["Inputs!B3"] == 1.2
