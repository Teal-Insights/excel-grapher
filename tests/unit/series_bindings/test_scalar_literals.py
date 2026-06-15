"""Unit tests for scalar Python literal rendering."""

from __future__ import annotations

from datetime import datetime

from excel_grapher.series_bindings.scalar_literals import py_scalar_literal


def test_py_scalar_literal_datetime() -> None:
    assert py_scalar_literal(datetime(2024, 1, 1)) == "datetime.datetime(2024, 1, 1, 0, 0)"
