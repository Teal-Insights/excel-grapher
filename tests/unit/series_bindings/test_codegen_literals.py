"""Unit tests for generated-code scalar literal helpers."""

from __future__ import annotations

from datetime import datetime
from typing import cast

from excel_grapher.series_bindings.codegen_literals import (
    emit_compute_preamble_lines,
    emit_setter_type_alias_lines,
    py_scalar_literal,
    python_annotation_for_dtype,
    resolution_includes_datetime,
    setter_input_annotation,
)
from excel_grapher.series_bindings.types import SeriesResolution


def test_py_scalar_literal_bool_and_datetime() -> None:
    assert py_scalar_literal(True) == "True"
    assert py_scalar_literal(False) == "False"
    assert py_scalar_literal(datetime(2024, 1, 1)) == "datetime.datetime(2024, 1, 1, 0, 0)"
    assert (
        py_scalar_literal(datetime(2024, 6, 15, 12, 30, 45, 123456))
        == "datetime.datetime(2024, 6, 15, 12, 30, 45, 123456)"
    )


def test_emit_setter_type_alias_lines_omit_datetime_when_unused() -> None:
    lines = emit_setter_type_alias_lines(include_datetime=False)
    assert lines[:2] == [
        "from collections.abc import Sequence",
        "from typing import TYPE_CHECKING, TypeAlias",
    ]
    assert "Scalar: TypeAlias = str | int | float | bool | None" in lines
    assert "if TYPE_CHECKING:" in lines
    assert any(
        line.strip() == "DataFrameInput: TypeAlias = pd.DataFrame | pl.DataFrame" for line in lines
    )
    assert any(line.strip() == "DataFrameInput: TypeAlias = object" for line in lines)
    assert "SeriesInput: TypeAlias = Records | Record | Sequence[Scalar] | DataFrameInput" in lines


def test_emit_setter_type_alias_lines_include_datetime_when_needed() -> None:
    lines = emit_setter_type_alias_lines(include_datetime=True)
    assert lines[0] == "from collections.abc import Sequence"
    assert "Scalar: TypeAlias = str | int | float | bool | datetime | None" in lines
    assert "SeriesInput: TypeAlias = Records | Record | Sequence[Scalar] | DataFrameInput" in lines


def test_emit_compute_preamble_lines_include_datetime_when_needed() -> None:
    lines = emit_compute_preamble_lines(include_datetime=True)
    assert lines[0] == "import datetime"
    assert lines[-1] == ""


def test_python_annotation_for_dtype() -> None:
    assert python_annotation_for_dtype("float") == "float"
    assert python_annotation_for_dtype("number") == "int | float"
    assert python_annotation_for_dtype("datetime") == "datetime"
    assert python_annotation_for_dtype("auto") is None
    assert python_annotation_for_dtype(None) is None


def test_setter_input_annotation_narrows_by_layout_and_dtype() -> None:
    assert (
        setter_input_annotation(
            layout="scalar",
            measure_dtype="float",
            scalar_shorthand=True,
        )
        == "Records | Record | float"
    )
    assert (
        setter_input_annotation(
            layout="series",
            measure_dtype="float",
            scalar_shorthand=False,
        )
        == "Records | Record | Sequence[float] | DataFrameInput"
    )
    assert (
        setter_input_annotation(
            layout="matrix",
            measure_dtype="float",
            scalar_shorthand=False,
        )
        == "Records | Record | DataFrameInput"
    )
    assert (
        setter_input_annotation(
            layout="series",
            measure_dtype=None,
            scalar_shorthand=False,
        )
        == "SeriesInput"
    )


def test_resolution_includes_datetime() -> None:
    resolved = {
        "series_id": "calendar",
        "ok": True,
        "requires_address": False,
        "leaves": [
            {
                "address": "Inputs!B2",
                "coordinates": {"TIME_PERIOD": datetime(2024, 1, 1)},
                "key": {"TIME_PERIOD": datetime(2024, 1, 1)},
                "record": {"TIME_PERIOD": datetime(2024, 1, 1), "OBS_VALUE": 1.0},
            }
        ],
        "issues": [],
    }
    assert resolution_includes_datetime(cast(SeriesResolution, resolved)) is True
