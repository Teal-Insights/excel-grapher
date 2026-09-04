"""Issue 687 — `as_measure` overloads narrow the default path to `float | str`."""

from __future__ import annotations

from datetime import datetime
from typing import Literal, get_overloads, get_type_hints

from excel_grapher.exporter.inverted_tree.runtime import as_measure


def test_as_measure_declares_literal_dtype_overloads() -> None:
    overloads = get_overloads(as_measure)
    assert overloads, "as_measure must declare @overload entries for Literal dtype"

    returns = {get_type_hints(fn)["return"] for fn in overloads}
    assert returns == {
        float | str,
        int | str,
        str,
        bool | str,
        datetime | str,
    }

    dtypes = {get_type_hints(fn)["dtype"] for fn in overloads}
    assert dtypes == {
        Literal["float"],
        Literal["int"],
        Literal["str"],
        Literal["bool"],
        Literal["datetime"],
    }
