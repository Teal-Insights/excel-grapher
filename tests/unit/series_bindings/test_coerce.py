"""Unit tests for series binding scalar coercion."""

from __future__ import annotations

from datetime import UTC, date, datetime

import pytest

from excel_grapher.series_bindings.coerce import (
    coerce_constant,
    coerce_scalar,
    validate_binding_scalar,
)

_EXCEL_EPOCH = datetime(1899, 12, 30)


def _excel_serial_for(when: datetime) -> float:
    """Build an Excel day-fraction serial for *when* relative to the 1899-12-30 epoch."""
    return (when - _EXCEL_EPOCH).total_seconds() / 86400


def test_coerce_scalar_auto_preserves_bool() -> None:
    assert coerce_scalar(True, "auto") is True
    assert coerce_scalar(False, "auto") is False


def test_coerce_scalar_bool_from_numeric() -> None:
    assert coerce_scalar(1, "bool") is True
    assert coerce_scalar(0, "bool") is False


def test_coerce_scalar_bool_from_float_zero_and_one() -> None:
    assert coerce_scalar(1.0, "bool") is True
    assert coerce_scalar(0.0, "bool") is False


@pytest.mark.parametrize("raw", [0.5, 2, -1, 1.5])
def test_coerce_scalar_bool_rejects_non_binary_numeric(raw: float | int) -> None:
    with pytest.raises(ValueError, match="Cannot coerce"):
        coerce_scalar(raw, "bool")


def test_coerce_scalar_bool_from_strings() -> None:
    assert coerce_scalar("TRUE", "bool") is True
    assert coerce_scalar("false", "bool") is False
    assert coerce_scalar("yes", "bool") is True
    assert coerce_scalar("no", "bool") is False


def test_coerce_scalar_bool_rejects_invalid() -> None:
    with pytest.raises(ValueError, match="Cannot coerce"):
        coerce_scalar("maybe", "bool")


def test_coerce_constant_bool_from_yaml_true() -> None:
    assert coerce_constant(True, read_as="auto") is True


def test_coerce_constant_bool_from_string_with_inferred_read() -> None:
    assert coerce_constant("TRUE", read_as="bool") is True


def test_coerce_scalar_auto_preserves_datetime() -> None:
    value = datetime(2024, 1, 15)
    assert coerce_scalar(value, "auto") == value


def test_coerce_scalar_auto_normalizes_date() -> None:
    assert coerce_scalar(date(2024, 1, 15), "auto") == datetime(2024, 1, 15)


def test_coerce_scalar_datetime_from_iso_string() -> None:
    assert coerce_scalar("2024-01-15", "datetime") == datetime(2024, 1, 15)
    assert coerce_scalar("2024-01-15T12:30:00", "datetime") == datetime(2024, 1, 15, 12, 30)


def test_coerce_scalar_datetime_from_excel_serial() -> None:
    expected = datetime(2024, 1, 1)
    serial = _excel_serial_for(expected)
    assert coerce_scalar(serial, "datetime") == expected


def test_coerce_scalar_datetime_from_excel_serial_noon() -> None:
    expected = datetime(2024, 1, 15, 12, 0, 0)
    serial = _excel_serial_for(expected)
    assert coerce_scalar(serial, "datetime") == expected


def test_coerce_scalar_datetime_from_excel_serial_preserves_milliseconds() -> None:
    expected = datetime(2024, 1, 15, 12, 30, 45, 500000)
    serial = _excel_serial_for(expected)
    assert coerce_scalar(serial, "datetime") == expected


def test_coerce_scalar_datetime_from_excel_serial_near_midnight() -> None:
    expected = datetime(2024, 1, 15, 23, 59, 59, 900000)
    serial = _excel_serial_for(expected)
    assert coerce_scalar(serial, "datetime") == expected


def test_coerce_scalar_datetime_rejects_timezone_aware() -> None:
    aware = datetime(2024, 1, 1, tzinfo=UTC)
    with pytest.raises(ValueError, match="Timezone-aware"):
        coerce_scalar(aware, "auto")
    with pytest.raises(ValueError, match="Timezone-aware"):
        coerce_scalar("2024-01-01T00:00:00Z", "datetime")


def test_coerce_constant_datetime_from_iso_string() -> None:
    assert coerce_constant("2024-06-30", read_as="datetime") == datetime(2024, 6, 30)


def test_coerce_module_does_not_depend_on_codegen_literals() -> None:
    from pathlib import Path

    source = Path("excel_grapher/series_bindings/coerce.py").read_text(encoding="utf-8")
    assert "codegen_literals" not in source
    from excel_grapher.series_bindings import coerce as coerce_module

    assert "py_scalar_literal" not in coerce_module.__all__


def test_validate_binding_scalar_float_accepts_int_and_float() -> None:
    assert validate_binding_scalar(1.5, "float") == 1.5
    assert validate_binding_scalar(2, "float") == 2.0


def test_validate_binding_scalar_float_rejects_string() -> None:
    with pytest.raises(TypeError, match="float"):
        validate_binding_scalar("not-a-number", "float")


def test_validate_binding_scalar_float_rejects_bool() -> None:
    with pytest.raises(TypeError, match="float"):
        validate_binding_scalar(True, "float")


def test_validate_binding_scalar_number_preserves_int() -> None:
    assert validate_binding_scalar(3, "number") == 3
    assert validate_binding_scalar(3.5, "number") == 3.5


def test_validate_binding_scalar_number_rejects_string() -> None:
    with pytest.raises(TypeError, match="number"):
        validate_binding_scalar("not-a-number", "number")


def test_validate_binding_scalar_int_rejects_float() -> None:
    with pytest.raises(TypeError, match="int"):
        validate_binding_scalar(4.0, "int")


def test_validate_binding_scalar_bool_rejects_int() -> None:
    with pytest.raises(TypeError, match="bool"):
        validate_binding_scalar(1, "bool")


def test_validate_binding_scalar_string_rejects_int() -> None:
    with pytest.raises(TypeError, match="string"):
        validate_binding_scalar(1, "string")


def test_validate_binding_scalar_datetime_accepts_date() -> None:
    assert validate_binding_scalar(date(2024, 1, 15), "datetime") == datetime(2024, 1, 15)


def test_validate_binding_scalar_datetime_rejects_string() -> None:
    with pytest.raises(TypeError, match="datetime"):
        validate_binding_scalar("2024-01-15", "datetime")


def test_validate_binding_scalar_none_passthrough() -> None:
    assert validate_binding_scalar(None, "float") is None


def test_validate_binding_scalar_auto_passthrough() -> None:
    assert validate_binding_scalar("x", "auto") == "x"
