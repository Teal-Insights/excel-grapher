"""Element-wise `IF` over scalars and aligned arrays (#732)."""

from __future__ import annotations

from typing import cast

from excel_grapher.core.grid import Range
from excel_grapher.core.logic_funcs import logical_if
from excel_grapher.core.types import CellValue, XlError


def test_logical_if_scalar_picks_a_branch() -> None:
    assert logical_if(True, 10, 20) == 10
    assert logical_if(False, 10, 20) == 20
    assert logical_if(0, 10, 20) == 20
    assert logical_if(0, 10) is False


def test_logical_if_omitted_else_is_false() -> None:
    assert logical_if([[True], [False]], [[1], [2]]) == [[1], [False]]


def test_logical_if_selects_aligned_then_else_elements() -> None:
    cond = [[True], [False], [True]]
    then = [[10], [20], [30]]
    otherwise = [[100], [200], [300]]
    assert logical_if(cond, then, otherwise) == [[10], [200], [30]]


def test_logical_if_broadcasts_scalar_else() -> None:
    cond = [[True, False], [False, True]]
    then = [[1, 2], [3, 4]]
    assert logical_if(cond, then, 0) == [[1, 0], [0, 4]]


def test_logical_if_shape_mismatch_is_value_error() -> None:
    assert logical_if([[True], [False]], [[1], [2], [3]], 0) == XlError.VALUE


def test_logical_if_over_lazy_range() -> None:
    cond = Range("S", 1, 1, 2, 1, lambda a: {"S!A1": True, "S!A2": False}[a])
    then = Range("S", 1, 2, 2, 2, lambda a: {"S!B1": 10, "S!B2": 20}[a])
    result = logical_if(cast(CellValue, cond), cast(CellValue, then), 0)
    assert result == [[10], [0]]
