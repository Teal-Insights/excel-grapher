"""Unit tests for generated-code scalar literal helpers."""

from __future__ import annotations

from datetime import datetime
from typing import cast

from excel_grapher.series_bindings.codegen_literals import (
    emit_compute_preamble_lines,
    emit_setter_type_alias_lines,
    py_scalar_literal,
    resolution_includes_datetime,
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
    assert lines == [
        "Scalar = str | int | float | bool | None",
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
    ]


def test_emit_setter_type_alias_lines_include_datetime_when_needed() -> None:
    lines = emit_setter_type_alias_lines(include_datetime=True)
    assert lines[0] == "import datetime"
    assert "datetime.datetime | None" in lines[2]


def test_emit_compute_preamble_lines_include_datetime_when_needed() -> None:
    lines = emit_compute_preamble_lines(include_datetime=True)
    assert lines[0] == "import datetime"
    assert lines[-1] == ""


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
