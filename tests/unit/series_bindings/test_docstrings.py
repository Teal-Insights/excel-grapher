"""Unit tests for series-binding docstring callbacks and contract derivation."""

from __future__ import annotations

from collections.abc import Mapping
from datetime import datetime
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
from excel_grapher.series_bindings.docstring_renderers import (
    GoogleSeriesDocstringRenderer,
    NumpySeriesDocstringRenderer,
    PlainSeriesDocstringRenderer,
    RstSeriesDocstringRenderer,
    _render_example_call,
    _render_positional_example,
    resolve_series_docstring_renderer,
)
from excel_grapher.series_bindings.docstrings import (
    FieldContract,
    FieldDoc,
    SeriesBindingDocstringContext,
    SeriesBindingDocstringContract,
    SeriesFunctionDoc,
    derive_doc_contract,
    emit_docstring_literal,
    list_series_docstring_callbacks,
    register_series_docstring_callback,
    resolve_series_docstring_callback,
    resolve_series_function_docstring,
    run_series_docstring_callback,
    unregister_series_docstring_callback,
)
from excel_grapher.series_bindings.setter_codegen import _canonical_key_order
from excel_grapher.series_bindings.types import SeriesResolution, WorkbookSeriesBindings
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


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
        callback_spec=name,
    )
    assert rendered is not None
    assert rendered.startswith("second")


def test_resolve_series_docstring_callback_returns_direct_callable() -> None:
    def direct(ctx: SeriesBindingDocstringContext) -> SeriesFunctionDoc:
        del ctx
        return SeriesFunctionDoc(
            summary="direct",
            purpose="direct",
            record_matching="direct",
        )

    resolved = resolve_series_docstring_callback(direct)
    assert resolved is direct


def test_resolve_series_docstring_callback_unknown_name_raises() -> None:
    with pytest.raises(ValueError, match="Unknown series docstring callback"):
        resolve_series_docstring_callback("not.registered.callback")


def test_unregister_series_docstring_callback_removes_registration() -> None:
    name = "_test_unregister_docstring_callback"

    def _noop(ctx: SeriesBindingDocstringContext) -> None:
        del ctx
        return None

    register_series_docstring_callback(name, _noop)
    assert name in list_series_docstring_callbacks()
    unregister_series_docstring_callback(name)
    assert name not in list_series_docstring_callbacks()


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
    assert contract.layout == "series"
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


def test_derive_doc_contract_reports_bool_dtype() -> None:
    contract = derive_doc_contract(
        {
            "id": "bool_flag",
            "data_range": "Flags!A2",
            "layout": "scalar",
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
        },
        function_kind="setter",
        function_name="set_bool_flag",
        resolution={
            "series_id": "bool_flag",
            "ok": True,
            "requires_address": False,
            "leaves": [
                {
                    "address": "Flags!A2",
                    "coordinates": {"IS_ACTIVE": True, "OBS_VALUE": False},
                    "key": {"IS_ACTIVE": True},
                    "record": {"IS_ACTIVE": True, "OBS_VALUE": False},
                }
            ],
            "issues": [],
        },
        bindings={
            "schema_version": "1.4.0",
            "workbook": "flags.xlsx",
            "series": [],
            "concept_scheme": {
                "id": "flags",
                "concepts": [{"id": "IS_ACTIVE", "dtype": "bool"}],
            },
        },
    )

    assert contract.value_type == "bool"
    assert contract.fields["OBS_VALUE"].dtype == "bool"
    assert contract.fields["IS_ACTIVE"].dtype == "bool"
    assert contract.example_records[0]["IS_ACTIVE"] is True
    assert contract.example_records[0]["OBS_VALUE"] is False


def test_derive_doc_contract_reports_datetime_dtype() -> None:
    period = datetime(2024, 1, 1)
    contract = derive_doc_contract(
        {
            "id": "calendar_series",
            "data_range": "Inputs!B2",
            "layout": "scalar",
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
        },
        function_kind="compute",
        function_name="compute_calendar_series",
        resolution={
            "series_id": "calendar_series",
            "ok": True,
            "requires_address": False,
            "leaves": [
                {
                    "address": "Inputs!B2",
                    "coordinates": {"TIME_PERIOD": period, "OBS_VALUE": 1.5},
                    "key": {"TIME_PERIOD": period},
                    "record": {"TIME_PERIOD": period, "OBS_VALUE": 1.5},
                }
            ],
            "issues": [],
        },
        bindings={
            "schema_version": "1.4.0",
            "workbook": "calendar.xlsx",
            "series": [],
            "concept_scheme": {
                "id": "calendar",
                "concepts": [{"id": "TIME_PERIOD", "dtype": "datetime"}],
            },
        },
    )

    assert contract.value_type == "float"
    assert contract.fields["TIME_PERIOD"].dtype == "datetime"
    assert contract.example_records[0]["TIME_PERIOD"] == period
    assert isinstance(contract.example_records[0]["OBS_VALUE"], float)


def test_derive_doc_contract_infers_dtype_from_dimension_bind_read() -> None:
    contract = derive_doc_contract(
        {
            "id": "calendar_series",
            "data_range": "Inputs!B2",
            "layout": "scalar",
            "structure": {
                "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                "dimensions": [
                    {
                        "concept": "TIME_PERIOD",
                        "bind": {
                            "kind": "column_header",
                            "header_row": 1,
                            "read": "datetime",
                        },
                    }
                ],
            },
            "key": ["TIME_PERIOD"],
        },
        function_kind="setter",
        function_name="set_calendar_series",
        resolution={
            "series_id": "calendar_series",
            "ok": True,
            "requires_address": False,
            "leaves": [],
            "issues": [],
        },
        bindings=_bindings_stub(),
    )

    assert contract.fields["TIME_PERIOD"].dtype == "datetime"


def _demo_contract_and_doc() -> tuple[SeriesBindingDocstringContract, SeriesFunctionDoc]:
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
    return contract, doc


def _series_contract_and_doc(
    *,
    layout: str = "series",
    keys: tuple[str, ...] = ("COUNTRY",),
    positional_example_values: tuple[object, ...] | None = (60.0, 80.0),
) -> tuple[SeriesBindingDocstringContract, SeriesFunctionDoc]:
    example_records: list[dict[str, object]] = [
        {**{key: country for key in keys}, "OBS_VALUE": value}
        for country, value in (("Borvelia", 60.0), ("Litellia", 80.0))
    ]
    if len(keys) != 1 or layout != "series":
        positional_example_values = None
    contract = SeriesBindingDocstringContract(
        series_id="country_initial_debt",
        function_name="set_country_initial_debt",
        function_kind="setter",
        data_range="Inputs!B10:B12",
        layout=layout,
        value_type="float",
        required_fields=(*keys, "OBS_VALUE"),
        fields={
            **{
                key: FieldContract(
                    concept_name=key,
                    dtype="string",
                    required=True,
                    expected_value=None,
                )
                for key in keys
            },
            "OBS_VALUE": FieldContract(
                concept_name="Observation value",
                dtype="float",
                required=True,
                expected_value=None,
            ),
        },
        example_records=tuple(example_records),
        notes="",
        positional_example_values=positional_example_values,
    )
    doc = SeriesFunctionDoc(
        summary="Set country initial debt.",
        purpose="Updates the bound range.",
        record_matching="Match by key.",
        field_descriptions={
            **{key: FieldDoc(description=f"Describe {key}.") for key in keys},
            "OBS_VALUE": FieldDoc(description="Observation value."),
        },
    )
    return contract, doc


def test_scalar_setter_docstring_describes_scalar_shapes() -> None:
    contract, doc = _demo_contract_and_doc()
    rendered = GoogleSeriesDocstringRenderer().render(contract, doc, series={"id": "demo"})
    assert "bare scalar" in rendered
    assert "tidy pandas/polars DataFrame" not in rendered
    assert "empty_measure" not in rendered
    assert "set_demo(ctx, 2.0)" in rendered
    assert "set_demo(ctx, [" not in rendered


def test_series_setter_docstring_describes_series_shapes() -> None:
    contract, doc = _series_contract_and_doc(layout="series", keys=("COUNTRY",))
    rendered = GoogleSeriesDocstringRenderer().render(
        contract, doc, series={"id": "country_initial_debt"}
    )
    assert "tidy pandas/polars DataFrame" in rendered
    assert "1-D iterable of measure values in key order" in rendered
    assert "empty_measure" in rendered
    assert "set_country_initial_debt(ctx, [" in rendered
    assert "set_country_initial_debt(ctx, [60.0, 80.0])" in rendered


def test_render_positional_example_uses_full_key_order_values() -> None:
    contract, _doc = _series_contract_and_doc(
        positional_example_values=(60.0, 80.0, 100.0),
    )
    assert _render_positional_example(contract) == [
        "set_country_initial_debt(ctx, [60.0, 80.0, 100.0])"
    ]


def test_render_positional_example_omitted_without_full_values() -> None:
    """Partial example_records must not advertise an invalid positional call."""
    contract, _doc = _series_contract_and_doc(positional_example_values=None)
    assert _render_positional_example(contract) == []
    example_call = "\n".join(_render_example_call(contract))
    assert "set_country_initial_debt(ctx, [" in example_call
    assert "set_country_initial_debt(ctx, [60.0, 80.0])" not in example_call


def test_derive_doc_contract_positional_values_match_key_order(tmp_path: Path) -> None:
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
    key_fields = [str(field) for field in (series.get("key") or [])]
    key_order = _canonical_key_order(resolved, key_fields)
    assert key_order is not None
    assert len(key_order) == 5

    contract = derive_doc_contract(
        series,
        function_kind="setter",
        function_name="set_borvelia_primary_balance",
        resolution=resolved,
        bindings=bindings,
    )

    assert len(contract.example_records) == 2
    assert contract.positional_example_values == (-2.0, -1.0, 0.0, 1.0, 2.0)
    assert len(contract.positional_example_values) == len(key_order)

    rendered = GoogleSeriesDocstringRenderer().render(
        contract,
        SeriesFunctionDoc(
            summary="Set borvelia_primary_balance.",
            purpose="Updates workbook inputs from Records.",
            record_matching="Records match cells by key fields.",
            field_descriptions={
                field: FieldDoc(description=f"Value for {field}.")
                for field in contract.required_fields
            },
        ),
        series=series,
    )
    assert "set_borvelia_primary_balance(ctx, [-2.0, -1.0, 0.0, 1.0, 2.0])" in rendered
    assert "set_borvelia_primary_balance(ctx, [-2.0, -1.0])" not in rendered


def test_series_multi_key_setter_omits_positional_shape() -> None:
    contract, doc = _series_contract_and_doc(layout="series", keys=("COUNTRY", "TIME_PERIOD"))
    rendered = GoogleSeriesDocstringRenderer().render(
        contract, doc, series={"id": "country_initial_debt"}
    )
    assert "tidy pandas/polars DataFrame" in rendered
    assert "1-D iterable of measure values in key order" not in rendered
    assert "empty_measure" in rendered


def test_matrix_setter_docstring_describes_matrix_shapes() -> None:
    contract, doc = _series_contract_and_doc(layout="matrix", keys=("INDICATOR", "TIME_PERIOD"))
    rendered = GoogleSeriesDocstringRenderer().render(
        contract, doc, series={"id": "country_initial_debt"}
    )
    assert "tidy pandas/polars DataFrame" in rendered
    assert "1-D iterable of measure values in key order" not in rendered
    assert "empty_measure" in rendered
    assert "set_country_initial_debt(ctx, [" in rendered


@pytest.mark.parametrize(
    "renderer_cls",
    [
        PlainSeriesDocstringRenderer,
        RstSeriesDocstringRenderer,
        GoogleSeriesDocstringRenderer,
        NumpySeriesDocstringRenderer,
    ],
)
def test_all_renderers_describe_layout_specific_input(renderer_cls: type) -> None:
    scalar_contract, doc = _demo_contract_and_doc()
    scalar_rendered = renderer_cls().render(scalar_contract, doc, series={"id": "demo"})
    assert "bare scalar" in scalar_rendered
    assert "empty_measure" not in scalar_rendered

    series_contract, series_doc = _series_contract_and_doc()
    series_rendered = renderer_cls().render(
        series_contract, series_doc, series={"id": "country_initial_debt"}
    )
    assert "tidy pandas/polars DataFrame" in series_rendered
    assert "empty_measure" in series_rendered


def test_resolve_series_docstring_renderer_builtin_names() -> None:
    for name in ("plain", "rst", "google", "numpy"):
        renderer = resolve_series_docstring_renderer(name)
        assert hasattr(renderer, "render")


def test_resolve_series_docstring_renderer_unknown_raises() -> None:
    with pytest.raises(ValueError, match="Unknown series docstring renderer"):
        resolve_series_docstring_renderer(cast(Any, "sphinx"))


def test_resolve_series_docstring_renderer_custom_object() -> None:
    class _CustomRenderer:
        def render(
            self,
            contract: SeriesBindingDocstringContract,
            doc: SeriesFunctionDoc,
            *,
            series: Mapping[str, Any] | None = None,
        ) -> str:
            del contract, series
            return f"CUSTOM:{doc.summary}"

    renderer = resolve_series_docstring_renderer(_CustomRenderer())
    contract, doc = _demo_contract_and_doc()
    assert renderer.render(contract, doc) == "CUSTOM:Set demo values."


def test_resolve_series_docstring_renderer_custom_callable() -> None:
    def _custom_callable(
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
        *,
        series: Any = None,
    ) -> str:
        del contract, series
        return f"CALLABLE:{doc.summary}"

    renderer = resolve_series_docstring_renderer(_custom_callable)
    contract, doc = _demo_contract_and_doc()
    assert renderer.render(contract, doc) == "CALLABLE:Set demo values."


def test_resolve_series_docstring_renderer_invalid_custom_raises() -> None:
    with pytest.raises(TypeError, match="Custom renderer must be"):
        resolve_series_docstring_renderer(cast(Any, 42))


@pytest.mark.parametrize(
    ("renderer_cls", "markers"),
    [
        (PlainSeriesDocstringRenderer, ("Required record fields:", "Source binding:", "Example:")),
        (RstSeriesDocstringRenderer, (":TIME_PERIOD:", "Source binding", "Example")),
        (GoogleSeriesDocstringRenderer, ("Args:", "Returns:", "Examples:")),
        (
            NumpySeriesDocstringRenderer,
            ("Parameters\n----------", "Returns\n-------", "Examples\n--------"),
        ),
    ],
)
def test_builtin_renderer_includes_sections(
    renderer_cls: type,
    markers: tuple[str, ...],
) -> None:
    contract, doc = _demo_contract_and_doc()
    rendered = renderer_cls().render(contract, doc, series={"id": "demo"})
    assert "Set demo values." in rendered
    assert "TIME_PERIOD" in rendered
    assert "set_demo(ctx, 2.0)" in rendered
    for marker in markers:
        assert marker in rendered


def test_plain_renderer_includes_sections() -> None:
    contract, doc = _demo_contract_and_doc()
    rendered = PlainSeriesDocstringRenderer().render(contract, doc, series={"id": "demo"})
    assert "Set demo values." in rendered
    assert "Required record fields:" in rendered
    assert "TIME_PERIOD:" in rendered
    assert "Source binding:" in rendered
    assert "Example:" in rendered
    assert "set_demo(ctx, 2.0)" in rendered


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
        callback_spec=None,
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
        callback_spec=callback_name,
    )
    assert doc is not None
    assert "Set borvelia_primary_balance." in doc
    assert "Required record fields:" in doc


def test_resolve_series_function_docstring_applies_renderer(tmp_path: Path) -> None:
    callback_name = "_test_renderer_selection_callback"
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
        callback_spec=callback_name,
        docstring_renderer="google",
    )
    assert doc is not None
    assert "Set borvelia_primary_balance." in doc
    assert "Args:" in doc
    assert "Required record fields:" in doc
