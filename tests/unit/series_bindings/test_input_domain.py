"""Schema, coercion, and codegen tests for `input.domain` (schema 1.13.0)."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict
from excel_grapher.series_bindings import (
    resolve_series_binding,
)
from excel_grapher.series_bindings.input_coerce import (
    coerce_setter_input,
    measure_domain_from_series,
    require_input_domain,
)
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    validate_bindings_document,
)
from excel_grapher.series_bindings.setter_codegen import (
    emit_input_coerce_helpers,
    emit_setter_function,
    emit_setter_helpers,
)
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS


def _scalar_string_doc(*, domain: dict[str, Any] | None = None) -> dict[str, Any]:
    input_block: dict[str, Any] = {"setter": {"name": "set_country"}}
    if domain is not None:
        input_block["domain"] = domain
    return {
        "schema_version": "1.13.0",
        "series": [
            {
                "id": "country",
                "sheet": "Dash",
                "data_range": "Dash!B1",
                "layout": "scalar",
                "input": input_block,
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "string",
                        "bind": {"kind": "data_cell", "read": "string"},
                    },
                    "dimensions": [],
                },
                "key": [],
            }
        ],
    }


def _scalar_float_doc(*, domain: dict[str, Any] | None = None) -> dict[str, Any]:
    input_block: dict[str, Any] = {"setter": {"name": "set_rate"}}
    if domain is not None:
        input_block["domain"] = domain
    return {
        "schema_version": "1.13.0",
        "series": [
            {
                "id": "rate",
                "sheet": "Dash",
                "data_range": "Dash!C1",
                "layout": "scalar",
                "input": input_block,
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


def _exec_setters(lines: list[str]) -> dict[str, object]:
    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
    }
    source = emit_input_coerce_helpers() + emit_setter_helpers() + lines
    exec("\n".join(source), namespace)
    return namespace


def test_schema_version_1_13_0_supported() -> None:
    assert "1.13.0" in SUPPORTED_SCHEMA_VERSIONS


def test_schema_accepts_enum_between_and_real_between_domains() -> None:
    for domain in (
        {"enum": ["Alpha", "Beta", "High "]},
        {"between": {"min": 0, "max": 100}},
        {"between": {"min": 0}},
        {"real_between": {"min": 0.0, "max": 300.0}},
        {"real_between": {"max": 1}},
    ):
        doc = (
            _scalar_string_doc(domain=domain)
            if "enum" in domain
            else _scalar_float_doc(domain=domain)
        )
        if "between" in domain:
            doc = {
                "schema_version": "1.13.0",
                "series": [
                    {
                        "id": "years",
                        "sheet": "Dash",
                        "data_range": "Dash!D1",
                        "layout": "scalar",
                        "input": {
                            "setter": {"name": "set_years"},
                            "domain": domain,
                        },
                        "structure": {
                            "measure": {
                                "concept": "OBS_VALUE",
                                "dtype": "int",
                                "bind": {"kind": "data_cell", "read": "int"},
                            },
                            "dimensions": [],
                        },
                        "key": [],
                    }
                ],
            }
        bindings = validate_bindings_document(doc)
        assert bindings["series"][0]["input"]["domain"] == domain


def test_schema_rejects_empty_enum_and_mixed_domain_keys() -> None:
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(_scalar_string_doc(domain={"enum": []}))
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(
            _scalar_float_doc(domain={"enum": ["a"], "real_between": {"min": 0}})
        )
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(_scalar_float_doc(domain={"real_between": {}}))


def test_coerce_enum_domain_rejects_and_accepts() -> None:
    domain = {"enum": frozenset({"Alpha", "Beta", "High "})}
    assert coerce_setter_input(
        "Beta",
        layout="scalar",
        key_fields=(),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
        measure_dtype="string",
        measure_domain=domain,
    ) == [{"OBS_VALUE": "Beta"}]
    assert coerce_setter_input(
        "High ",
        layout="scalar",
        key_fields=(),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
        measure_dtype="string",
        measure_domain=domain,
    ) == [{"OBS_VALUE": "High "}]
    with pytest.raises(ValueError, match="out of domain"):
        coerce_setter_input(
            "Nonexistent",
            layout="scalar",
            key_fields=(),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
            measure_dtype="string",
            measure_domain=domain,
        )
    with pytest.raises(ValueError, match="out of domain"):
        coerce_setter_input(
            "High",
            layout="scalar",
            key_fields=(),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
            measure_dtype="string",
            measure_domain=domain,
        )


def test_coerce_real_between_domain_inclusive_bounds() -> None:
    domain = {"real_between": {"min": 0.0, "max": 300.0}}
    assert coerce_setter_input(
        0,
        layout="scalar",
        key_fields=(),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
        measure_dtype="float",
        measure_domain=domain,
    ) == [{"OBS_VALUE": 0.0}]
    assert coerce_setter_input(
        300,
        layout="scalar",
        key_fields=(),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
        measure_dtype="float",
        measure_domain=domain,
    ) == [{"OBS_VALUE": 300.0}]
    with pytest.raises(ValueError, match="out of domain"):
        coerce_setter_input(
            301,
            layout="scalar",
            key_fields=(),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
            measure_dtype="float",
            measure_domain=domain,
        )


def test_coerce_between_domain_rejects_out_of_range_int() -> None:
    domain = {"between": {"min": 0, "max": 100}}
    with pytest.raises(ValueError, match="out of domain"):
        coerce_setter_input(
            -1,
            layout="scalar",
            key_fields=(),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
            measure_dtype="int",
            measure_domain=domain,
        )


def test_coerce_series_domain_checks_each_record() -> None:
    domain = {"real_between": {"min": -20.0, "max": 20.0}}
    with pytest.raises(ValueError, match=r"record\[1\].*out of domain"):
        coerce_setter_input(
            [
                {"TIME_PERIOD": 1, "OBS_VALUE": 1.0},
                {"TIME_PERIOD": 2, "OBS_VALUE": 21.0},
            ],
            layout="series",
            key_fields=("TIME_PERIOD",),
            measure_field="OBS_VALUE",
            key_order=(1, 2),
            strict=True,
            measure_dtype="float",
            measure_domain=domain,
        )


def test_emit_scalar_setter_rejects_out_of_enum_domain(tmp_path: Path) -> None:
    wb_path = tmp_path / "country.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Dash")
    ws.write("B1", "Alpha")
    wb.close()

    doc = _scalar_string_doc(domain={"enum": ["Alpha", "Beta", "High "]})
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]
    graph = create_dependency_graph(wb_path, ["Dash!B1"], load_values=True)
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved, bindings=bindings)
    code = "\n".join(lines)
    assert "measure_domain=" in code

    ns = _exec_setters(lines)
    setter = cast(Callable[..., None], ns["set_country"])
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, "Beta")
    assert ctx.inputs["Dash!B1"] == "Beta"
    setter(ctx, "High ")
    assert ctx.inputs["Dash!B1"] == "High "
    with pytest.raises(ValueError, match="out of domain"):
        setter(ctx, "Nonexistent")


def test_emit_scalar_setter_rejects_out_of_real_between_domain(tmp_path: Path) -> None:
    wb_path = tmp_path / "rate.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Dash")
    ws.write_number("C1", 1.0)
    wb.close()

    doc = _scalar_float_doc(domain={"real_between": {"min": 0, "max": 300}})
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]
    graph = create_dependency_graph(wb_path, ["Dash!C1"], load_values=True)
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved, bindings=bindings)
    ns = _exec_setters(lines)
    setter = cast(Callable[..., None], ns["set_rate"])
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx, 0)
    setter(ctx, 300)
    with pytest.raises(ValueError, match="out of domain"):
        setter(ctx, 301)


def test_measure_domain_from_series_normalizes_enum() -> None:
    series = _scalar_string_doc(domain={"enum": ["Alpha", "Beta"]})["series"][0]
    assert measure_domain_from_series(series) == {"enum": frozenset({"Alpha", "Beta"})}
    assert measure_domain_from_series(_scalar_string_doc()["series"][0]) is None


def test_require_input_domain_scalar_and_sequence() -> None:
    enum_domain = {"enum": frozenset({0, 1})}
    require_input_domain(0, enum_domain, series_id="flag")
    require_input_domain(1, enum_domain, series_id="flag")
    with pytest.raises(ValueError, match=r"flag out of domain: 2 not in \{0, 1\}"):
        require_input_domain(2, enum_domain, series_id="flag")

    bounds = {"real_between": {"min": 0, "max": 1}}
    require_input_domain((0.0, 1.0), bounds, series_id="rate")
    with pytest.raises(ValueError, match=r"rate\[1\] out of domain"):
        require_input_domain((0.0, 1.1), bounds, series_id="rate")
    with pytest.raises(ValueError, match=r"not in real_between"):
        require_input_domain((0.0, 1.1), bounds, series_id="rate")
