from __future__ import annotations

import pytest

from excel_grapher.core import CellValue, XlError
from excel_grapher.exporter.export_runtime import Range, XlErrorException
from excel_grapher.exporter.export_runtime.error_funcs import xl_isnumber, xl_istext
from excel_grapher.exporter.export_runtime.errors import xl_raise


def test_xl_error_exception_carries_excel_error_code() -> None:
    err = XlErrorException(XlError.DIV)

    assert err.code == XlError.DIV
    assert str(err) == "#DIV/0!"


def test_range_cell_access_evaluates_only_requested_cell() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": 1, "S!B1": 2, "S!A2": 3, "S!B2": 4}[address]

    rng = Range("S", 1, 1, 2, 2, resolve)

    assert rng.shape == (2, 2)
    assert rng.cell(1, 2) == 2
    assert calls == ["S!B1"]


def test_range_iteration_is_row_major() -> None:
    calls: list[str] = []
    values: dict[str, CellValue] = {"S!A1": 1, "S!B1": 2, "S!A2": 3, "S!B2": 4}

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return values[address]

    rng = Range("S", 1, 1, 2, 2, resolve)

    assert list(rng) == [1, 2, 3, 4]
    assert calls == ["S!A1", "S!B1", "S!A2", "S!B2"]


def test_range_iteration_raises_first_excel_error() -> None:
    calls: list[str] = []
    values: dict[str, CellValue] = {
        "S!A1": 1,
        "S!B1": XlError.NA,
        "S!A2": XlError.DIV,
        "S!B2": 4,
    }

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return values[address]

    rng = Range("S", 1, 1, 2, 2, resolve)

    with pytest.raises(XlErrorException) as exc_info:
        list(rng)

    assert exc_info.value.code == XlError.NA
    assert calls == ["S!A1", "S!B1"]


def test_range_row_and_column_views_preserve_laziness() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return address

    rng = Range("S", 1, 1, 3, 3, resolve)

    assert rng.row(2).shape == (1, 3)
    assert rng.row(2).cell(1, 3) == "S!C2"
    assert rng.column(1).shape == (3, 1)
    assert rng.column(1).cell(3, 1) == "S!A3"
    assert calls == ["S!C2", "S!A3"]


def test_xl_isnumber_istext_thunks_return_false_for_errors() -> None:
    """Thunked ISNUMBER/ISTEXT catch raised errors and reject XlError sentinels."""
    assert xl_isnumber(lambda: xl_raise(XlError.DIV)) is False
    assert xl_isnumber(lambda: XlError.NA) is False
    assert xl_isnumber(lambda: 2) is True
    assert xl_istext(lambda: xl_raise(XlError.NA)) is False
    assert xl_istext(lambda: XlError.DIV) is False
    assert xl_istext(lambda: "hello") is True
    assert xl_istext(lambda: 123) is False
