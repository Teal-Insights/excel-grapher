"""Evaluator ``xl_sumproduct`` wrapper materializes large grids for NumPy fastpath."""

from __future__ import annotations

from typing import cast

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.grid import Range
from excel_grapher.core.operators_fastpath import MIN_OPERATOR_FASTPATH_CELLS
from excel_grapher.core.sumproduct import sumproduct_cells
from excel_grapher.runtime.math import xl_sumproduct


def test_xl_sumproduct_small_delegates_to_sumproduct_cells() -> None:
    left = Range("S", 1, 1, 3, 1, lambda a: {"S!A1": 1, "S!A2": 2, "S!A3": 3}[a])
    right = Range("S", 1, 2, 3, 2, lambda a: {"S!B1": 4, "S!B2": 5, "S!B3": 6}[a])
    assert xl_sumproduct(cast(CellValue, left), cast(CellValue, right)) == 32.0
    assert sumproduct_cells(cast(CellValue, left), cast(CellValue, right)) == 32.0


def test_xl_sumproduct_large_uses_fastpath_when_available(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    nrows = MIN_OPERATOR_FASTPATH_CELLS
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        row = int(address.split("!")[1][1:])
        return float(row)

    seen_fastpath: list[int] = []

    def _fake_fastpath(arrays: list[np.ndarray]) -> float:
        seen_fastpath.append(arrays[0].size)
        return 42.0

    monkeypatch.setattr(
        "excel_grapher.runtime.math.try_fastpath_sumproduct",
        _fake_fastpath,
    )

    left = Range("S", 1, 1, nrows, 1, resolve)
    right = Range("S", 1, 2, nrows, 2, lambda a: 2.0)
    assert xl_sumproduct(cast(CellValue, left), cast(CellValue, right)) == 42.0
    assert seen_fastpath == [nrows]
    assert len(calls) == nrows


def test_xl_sumproduct_large_falls_back_to_reference_on_fastpath_miss(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    nrows = MIN_OPERATOR_FASTPATH_CELLS

    monkeypatch.setattr(
        "excel_grapher.runtime.math.try_fastpath_sumproduct",
        lambda *_args, **_kwargs: None,
    )

    left = np.arange(1, nrows + 1, dtype=object).reshape(nrows, 1)
    right = np.full((nrows, 1), 2.0, dtype=object)
    expected = float(sum(float(i) * 2.0 for i in range(1, nrows + 1)))
    assert xl_sumproduct(left, right) == expected


def test_xl_sumproduct_shape_mismatch_and_errors() -> None:
    assert (
        xl_sumproduct(
            np.array([[1.0], [2.0]], dtype=object),
            np.array([[1.0], [2.0], [3.0]], dtype=object),
        )
        == XlError.VALUE
    )
    assert (
        xl_sumproduct(
            np.array([[1.0], [XlError.NA]], dtype=object),
            np.array([[2.0], [2.0]], dtype=object),
        )
        == XlError.NA
    )
