"""Unit tests for series binding scalar coercion."""

from __future__ import annotations

from datetime import UTC, date, datetime

import pytest

from excel_grapher.series_bindings.coerce import (
    coerce_constant,
    coerce_scalar,
    py_scalar_literal,
)


def test_coerce_scalar_auto_preserves_bool() -> None:
    assert coerce_scalar(True, "auto") is True
    assert coerce_scalar(False, "auto") is False


def test_coerce_scalar_bool_from_numeric() -> None:
    assert coerce_scalar(1, "bool") is True
    assert coerce_scalar(0, "bool") is False


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
    serial = float((expected - datetime(1899, 12, 30)).days)
    assert coerce_scalar(serial, "datetime") == expected


def test_coerce_scalar_datetime_rejects_timezone_aware() -> None:
    aware = datetime(2024, 1, 1, tzinfo=UTC)
    with pytest.raises(ValueError, match="Timezone-aware"):
        coerce_scalar(aware, "auto")
    with pytest.raises(ValueError, match="Timezone-aware"):
        coerce_scalar("2024-01-01T00:00:00Z", "datetime")


def test_coerce_constant_datetime_from_iso_string() -> None:
    assert coerce_constant("2024-06-30", read_as="datetime") == datetime(2024, 6, 30)


def test_py_scalar_literal_datetime() -> None:
    assert py_scalar_literal(datetime(2024, 1, 1)) == "datetime.datetime(2024, 1, 1, 0, 0)"
