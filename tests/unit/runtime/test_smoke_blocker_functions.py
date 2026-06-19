"""Unit tests for smoke-blocker runtime functions."""

from __future__ import annotations

from datetime import datetime

import numpy as np
import pytest

from excel_grapher.core.coercions import datetime_to_excel_serial, to_number
from excel_grapher.core.types import XlError
from excel_grapher.runtime.datetime import xl_today
from excel_grapher.runtime.info import xl_iserror, xl_isna
from excel_grapher.runtime.math import xl_averageif, xl_countif
from excel_grapher.runtime.text import xl_lower, xl_value


def test_xl_iserror_and_isna() -> None:
    assert xl_iserror(XlError.DIV) is True
    assert xl_iserror(1) is False
    assert xl_isna(XlError.NA) is True
    assert xl_isna(XlError.DIV) is False


def test_xl_lower_and_value() -> None:
    assert xl_lower("AbC") == "abc"
    assert xl_value("12.5") == 12.5
    assert xl_value("not-a-number") == XlError.VALUE


def test_to_number_parses_iso_date_strings() -> None:
    serial = to_number("2018-03-15")
    assert not isinstance(serial, XlError)
    assert serial == pytest.approx(datetime_to_excel_serial(datetime(2018, 3, 15)))


def test_xl_today_returns_positive_serial() -> None:
    today = xl_today()
    assert today > 40_000


def test_xl_averageif_matches_category() -> None:
    categories = np.array([["A"], ["B"], ["A"]], dtype=object)
    values = np.array([[10.0], [20.0], [30.0]], dtype=object)
    assert xl_averageif(categories, "A", values) == 20.0
    assert xl_countif(categories, "A") == 2


def test_xl_averageif_no_matches_returns_div_error() -> None:
    categories = np.array([["A"]], dtype=object)
    values = np.array([[1.0]], dtype=object)
    assert xl_averageif(categories, "Z", values) == XlError.DIV
