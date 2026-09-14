"""Schema, coercion, and codegen tests for `input.domain` (schema 1.13.0)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings.input_coerce import (
    coerce_setter_input,
    measure_domain_from_series,
    require_input_domain,
)
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    validate_bindings_document,
)
from excel_grapher.series_bindings.types import ValidationReport
from excel_grapher.series_bindings.validate import validate_series_bindings
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS


def _scalar_string_doc(*, domain: dict[str, Any] | None = None) -> dict[str, Any]:
    input_block: dict[str, Any] = {}
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
    input_block: dict[str, Any] = {}
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


def _scalar_numeric_doc(
    *,
    dtype: str,
    domain: dict[str, Any],
    series_id: str = "share",
) -> dict[str, Any]:
    read = "float" if dtype == "number" else dtype
    return {
        "schema_version": "1.13.0",
        "workbook": "workbook.xlsx",
        "series": [
            {
                "id": series_id,
                "sheet": "Inputs",
                "data_range": "Inputs!A1",
                "layout": "scalar",
                "input": {"domain": domain},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": dtype,
                        "bind": {"kind": "data_cell", "read": read},
                    },
                    "dimensions": [],
                },
                "key": [],
            }
        ],
    }


def _write_scalar_input_workbook(path: Path, value: float | int = 0.0) -> Path:
    from fastpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    ws["A1"] = value
    wb.save(path)
    return path


def _report_for_domain_dtype(
    tmp_path: Path,
    *,
    dtype: str,
    domain: dict[str, Any],
    value: float | int = 0.0,
) -> ValidationReport:
    workbook = _write_scalar_input_workbook(tmp_path / "workbook.xlsx", value)
    graph = create_dependency_graph(workbook, ["Inputs!A1"], load_values=True)
    bindings = validate_bindings_document(_scalar_numeric_doc(dtype=dtype, domain=domain))
    return validate_series_bindings(graph, bindings, workbook=workbook)


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


def test_measure_domain_from_series_normalizes_enum() -> None:
    series = _scalar_string_doc(domain={"enum": ["Alpha", "Beta"]})["series"][0]
    assert measure_domain_from_series(series) == {"enum": frozenset({"Alpha", "Beta"})}
    assert measure_domain_from_series(_scalar_string_doc()["series"][0]) is None


def test_validate_rejects_float_dtype_with_between_domain(tmp_path: Path) -> None:
    report = _report_for_domain_dtype(
        tmp_path,
        dtype="float",
        domain={"between": {"min": 0, "max": 1}},
    )
    assert report["ok"] is False
    mismatches = [issue for issue in report["issues"] if issue["code"] == "domain_dtype_mismatch"]
    assert len(mismatches) == 1
    assert mismatches[0]["series_id"] == "share"
    assert "between" in mismatches[0]["message"]
    assert "float" in mismatches[0]["message"]
    assert "real_between" in mismatches[0]["message"]


def test_validate_rejects_number_dtype_with_between_domain(tmp_path: Path) -> None:
    report = _report_for_domain_dtype(
        tmp_path,
        dtype="number",
        domain={"between": {"min": 0, "max": 1}},
    )
    assert report["ok"] is False
    assert any(issue["code"] == "domain_dtype_mismatch" for issue in report["issues"])


def test_validate_rejects_int_dtype_with_real_between_domain(tmp_path: Path) -> None:
    report = _report_for_domain_dtype(
        tmp_path,
        dtype="int",
        domain={"real_between": {"min": 0, "max": 1}},
        value=0,
    )
    assert report["ok"] is False
    mismatches = [issue for issue in report["issues"] if issue["code"] == "domain_dtype_mismatch"]
    assert len(mismatches) == 1
    assert "real_between" in mismatches[0]["message"]
    assert "int" in mismatches[0]["message"]
    assert "between" in mismatches[0]["message"]


def test_validate_accepts_int_between_and_float_real_between(tmp_path: Path) -> None:
    int_report = _report_for_domain_dtype(
        tmp_path,
        dtype="int",
        domain={"between": {"min": 0, "max": 1}},
        value=0,
    )
    assert int_report["ok"] is True
    assert not any(issue["code"] == "domain_dtype_mismatch" for issue in int_report["issues"])

    float_report = _report_for_domain_dtype(
        tmp_path,
        dtype="float",
        domain={"real_between": {"min": 0, "max": 1}},
    )
    assert float_report["ok"] is True
    assert not any(issue["code"] == "domain_dtype_mismatch" for issue in float_report["issues"])

    number_report = _report_for_domain_dtype(
        tmp_path,
        dtype="number",
        domain={"real_between": {"min": 0, "max": 1}},
    )
    assert number_report["ok"] is True
    assert not any(issue["code"] == "domain_dtype_mismatch" for issue in number_report["issues"])


def test_validate_rejects_omitted_dtype_with_between_domain(tmp_path: Path) -> None:
    workbook = _write_scalar_input_workbook(tmp_path / "workbook.xlsx")
    graph = create_dependency_graph(workbook, ["Inputs!A1"], load_values=True)
    doc = _scalar_numeric_doc(dtype="float", domain={"between": {"min": 0, "max": 1}})
    measure = doc["series"][0]["structure"]["measure"]
    del measure["dtype"]
    measure["bind"] = {"kind": "data_cell"}
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is False
    mismatches = [issue for issue in report["issues"] if issue["code"] == "domain_dtype_mismatch"]
    assert len(mismatches) == 1
    assert "omitted" in mismatches[0]["message"]


def test_require_input_domain_between_names_wrong_type() -> None:
    domain = {"between": {"min": 0, "max": 1}}
    require_input_domain(0, domain, series_id="share")
    with pytest.raises(ValueError, match=r"share has type float; between requires int"):
        require_input_domain(0.0, domain, series_id="share")
    with pytest.raises(ValueError, match=r"not in between"):
        require_input_domain(2, domain, series_id="share")


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
