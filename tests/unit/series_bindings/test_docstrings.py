"""Unit tests for series-binding docstring callbacks and contract derivation."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)
from excel_grapher.series_bindings.docstrings import (
    FieldDoc,
    SeriesBindingDocstringContext,
    SeriesFunctionDoc,
    derive_doc_contract,
    emit_docstring_literal,
    list_series_docstring_callbacks,
    register_series_docstring_callback,
    render_series_function_doc,
    resolve_series_function_docstring,
    run_series_docstring_callback,
)
from excel_grapher.series_bindings.types import SeriesResolution, WorkbookSeriesBindings

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def _bindings_stub() -> WorkbookSeriesBindings:
    return {
        "schema_version": "1.0.0",
        "workbook": "unused.xlsx",
        "series": [],
        "concept_scheme": {},
    }


def _write_borvelia_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year - 3))
    wb.close()


def test_list_series_docstring_callbacks_includes_builtins() -> None:
    names = set(list_series_docstring_callbacks())
    assert "series_notes" in names


def test_register_series_docstring_callback_rejects_duplicate() -> None:
    def _noop(ctx: SeriesBindingDocstringContext) -> SeriesFunctionDoc | None:
        del ctx
        return None

    unique_name = "_test_duplicate_docstring_callback"
    register_series_docstring_callback(unique_name, _noop)
    with pytest.raises(ValueError, match="duplicate"):
        register_series_docstring_callback(unique_name, _noop)


def test_register_series_docstring_callback_allows_replace() -> None:
    name = "_test_replace_docstring_callback"
    register_series_docstring_callback(
        name,
        lambda ctx: SeriesFunctionDoc(
            summary="first",
            purpose="first",
            record_matching="first",
        ),
    )
    register_series_docstring_callback(
        name,
        lambda ctx: SeriesFunctionDoc(
            summary="second",
            purpose="second",
            record_matching="second",
        ),
        replace=True,
    )
    rendered = resolve_series_function_docstring(
        graph=DependencyGraph(),
        workbook=Path("unused.xlsx"),
        bindings=_bindings_stub(),
        series={
            "id": "demo",
            "data_range": "Sheet1!A1",
            "layout": "scalar",
            "structure": {"measure": {"concept": "OBS_VALUE"}},
            "key": [],
        },
        resolution={
            "series_id": "demo",
            "ok": True,
            "requires_address": False,
            "leaves": [
                {
                    "address": "S!A1",
                    "coordinates": {},
                    "key": {},
                    "record": {"OBS_VALUE": 1.0},
                }
            ],
            "issues": [],
        },
        function_kind="setter",
        function_name="set_demo",
        callback_name=name,
    )
    assert rendered is not None
    assert rendered.startswith("second")


def test_run_series_docstring_callback_unknown_name_raises() -> None:
    with pytest.raises(ValueError, match="Unknown series docstring callback"):
        run_series_docstring_callback("not.registered.callback", _make_context())


def _make_context() -> SeriesBindingDocstringContext:
    return SeriesBindingDocstringContext(
        graph=DependencyGraph(),
        workbook=Path("unused.xlsx"),
        bindings=_bindings_stub(),
        series={"id": "demo"},
        resolution={
            "series_id": "demo",
            "ok": True,
            "requires_address": False,
            "leaves": [],
            "issues": [],
        },
        contract=derive_doc_contract(
            {"id": "demo", "structure": {"measure": {"concept": "OBS_VALUE", "dtype": "float"}}},
            function_kind="setter",
            function_name="set_demo",
            resolution={
                "series_id": "demo",
                "ok": True,
                "requires_address": False,
                "leaves": [],
                "issues": [],
            },
            bindings=_bindings_stub(),
        ),
        function_kind="setter",
        function_name="set_demo",
    )


def test_derive_doc_contract_required_and_optional_fields(tmp_path: Path) -> None:
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

    contract = derive_doc_contract(
        series,
        function_kind="setter",
        function_name="set_borvelia_primary_balance",
        resolution=resolved,
        bindings=bindings,
    )

    assert contract.series_id == "borvelia_primary_balance"
    assert contract.function_name == "set_borvelia_primary_balance"
    assert contract.function_kind == "setter"
    assert contract.data_range == "Inputs!F5:J5"
    assert contract.layout == "row_series"
    assert contract.value_type == "float"
    assert contract.required_fields == ("TIME_PERIOD", "OBS_VALUE")
    assert contract.fields["TIME_PERIOD"].required is True
    assert contract.fields["UNIT_MEASURE"].required is False
    assert contract.fields["UNIT_MEASURE"].expected_value == "PC_GDP"
    assert len(contract.example_records) >= 1
    assert "TIME_PERIOD" in contract.example_records[0]
    assert "OBS_VALUE" in contract.example_records[0]


def test_derive_doc_contract_uses_none_for_unknown_dtype() -> None:
    contract = derive_doc_contract(
        {
            "id": "demo",
            "data_range": "Sheet1!A1",
            "layout": "scalar",
            "structure": {
                "measure": {"concept": "OBS_VALUE"},
                "dimensions": [{"concept": "TIME_PERIOD"}],
            },
            "key": ["TIME_PERIOD"],
        },
        function_kind="setter",
        function_name="set_demo",
        resolution={
            "series_id": "demo",
            "ok": True,
            "requires_address": False,
            "leaves": [
                {
                    "address": "Sheet1!A1",
                    "coordinates": {},
                    "key": {"TIME_PERIOD": 1},
                    "record": {"TIME_PERIOD": 1, "OBS_VALUE": 2.0},
                }
            ],
            "issues": [],
        },
        bindings={
            "schema_version": "1.0.0",
            "workbook": "demo.xlsx",
            "series": [],
            "concept_scheme": {},
        },
    )
    assert contract.value_type is None
    assert contract.fields["OBS_VALUE"].dtype is None
    assert contract.fields["TIME_PERIOD"].dtype is None


def test_render_series_function_doc_includes_sections() -> None:
    contract = derive_doc_contract(
        {
            "id": "demo",
            "data_range": "Sheet1!A1",
            "layout": "scalar",
            "structure": {
                "measure": {"concept": "OBS_VALUE", "dtype": "float"},
                "dimensions": [
                    {
                        "concept": "TIME_PERIOD",
                        "include_in_record": True,
                    }
                ],
            },
            "key": ["TIME_PERIOD"],
        },
        function_kind="setter",
        function_name="set_demo",
        resolution={
            "series_id": "demo",
            "ok": True,
            "requires_address": False,
            "leaves": [
                {
                    "address": "Sheet1!A1",
                    "coordinates": {},
                    "key": {"TIME_PERIOD": 1},
                    "record": {"TIME_PERIOD": 1, "OBS_VALUE": 2.0},
                }
            ],
            "issues": [],
        },
        bindings={
            "schema_version": "1.0.0",
            "workbook": "demo.xlsx",
            "series": [],
            "concept_scheme": {},
        },
    )
    doc = SeriesFunctionDoc(
        summary="Set demo values.",
        purpose="Updates the demo input series.",
        record_matching="Records match by TIME_PERIOD.",
        field_descriptions={
            "TIME_PERIOD": FieldDoc(description="Reporting period."),
            "OBS_VALUE": FieldDoc(description="Observed value."),
        },
    )
    rendered = render_series_function_doc(doc, contract=contract, series={"id": "demo"})
    assert "Set demo values." in rendered
    assert "Required record fields:" in rendered
    assert "TIME_PERIOD:" in rendered
    assert "Source binding:" in rendered
    assert "Example:" in rendered
    assert "set_demo(ctx, [" in rendered


def test_emit_docstring_literal_escapes_quotes_and_is_valid_python() -> None:
    doc = 'Contains "quotes" and\nmultiple lines.'
    lines = emit_docstring_literal(doc)
    code = "def fn():\n" + "\n".join(lines) + "\n    pass"
    exec(code, {})


def test_resolve_series_function_docstring_default_uses_series_notes() -> None:
    series: dict[str, Any] = {
        "id": "demo",
        "notes": "Custom notes for demo.",
        "structure": {"measure": {"concept": "OBS_VALUE", "dtype": "float"}},
        "key": [],
    }
    resolved = cast(
        SeriesResolution,
        {
            "series_id": "demo",
            "ok": True,
            "requires_address": False,
            "leaves": [{"address": "S!A1", "coordinates": {}, "key": {}, "record": {}}],
            "issues": [],
        },
    )
    doc = resolve_series_function_docstring(
        graph=DependencyGraph(),
        workbook=Path("unused.xlsx"),
        bindings=_bindings_stub(),
        series=series,
        resolution=resolved,
        function_kind="setter",
        function_name="set_demo",
        callback_name=None,
    )
    assert doc == "Custom notes for demo."


def test_resolve_series_function_docstring_registered_callback(tmp_path: Path) -> None:
    callback_name = "_test_structured_docstring_callback"
    register_series_docstring_callback(
        callback_name,
        lambda ctx: SeriesFunctionDoc(
            summary=f"Set {ctx.contract.series_id}.",
            purpose="Purpose text.",
            record_matching="Match by key.",
            field_descriptions={
                field: FieldDoc(description=f"Describe {field}.")
                for field in ctx.contract.required_fields
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

    doc = resolve_series_function_docstring(
        graph=graph,
        workbook=wb_path,
        bindings=bindings,
        series=series,
        resolution=resolved,
        function_kind="setter",
        function_name="set_borvelia_primary_balance",
        callback_name=callback_name,
    )
    assert doc is not None
    assert "Set borvelia_primary_balance." in doc
    assert "Required record fields:" in doc
