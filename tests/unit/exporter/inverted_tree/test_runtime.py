"""Runtime primitives for inverted-tree codegen (`take`, `require_aligned`, `xl_*`)."""

from __future__ import annotations

import pytest

from excel_grapher.exporter.inverted_tree.runtime import (
    XlError,
    require_aligned,
    require_length,
    take,
    xl_at,
    xl_choose,
    xl_div,
    xl_match,
    xl_raise,
)


def test_require_aligned_returns_common_length() -> None:
    assert require_aligned((1, 2, 3), ("a", "b", "c")) == 3


def test_require_aligned_rejects_mismatched_lengths() -> None:
    with pytest.raises(ValueError, match="misaligned"):
        require_aligned((1, 2), (1, 2, 3))


def test_require_aligned_rejects_empty_call() -> None:
    with pytest.raises(ValueError, match="at least one"):
        require_aligned()


def test_require_length_accepts_catalog_size() -> None:
    require_length((1, 2, 3), 3)


def test_require_length_rejects_mismatch() -> None:
    with pytest.raises(ValueError, match="expected length 3"):
        require_length((1, 2), 3)


def test_take_gathers_by_index() -> None:
    assert take((10, 20, 30, 40, 50), (0, 1, 2)) == (10, 20, 30)
    assert take((10, 20, 30, 40, 50), (1, 3)) == (20, 40)
    assert take((10, 20, 30), (2,)) == (30,)


def test_take_rejects_out_of_range() -> None:
    with pytest.raises(ValueError, match="take index 3"):
        take((1, 2, 3), (3,))
    with pytest.raises(ValueError, match="take index -1"):
        take((1, 2, 3), (-1,))


def test_xl_div_and_div_zero() -> None:
    assert xl_div(10.0, 2.0) == 5.0
    with pytest.raises(XlError) as exc:
        xl_div(1.0, 0.0)
    assert exc.value.code == "#DIV/0!"


def test_xl_choose_and_out_of_range() -> None:
    assert xl_choose(2, 10.0, 20.0, 30.0) == 20.0
    with pytest.raises(XlError) as exc:
        xl_choose(0, 1.0)
    assert exc.value.code == "#VALUE!"


def test_xl_match_exact_and_na() -> None:
    assert xl_match("Litellia", ("Borvelia", "Litellia", "Aurelium"), 0) == 2
    with pytest.raises(XlError) as exc:
        xl_match("Nope", ("Borvelia",), 0)
    assert exc.value.code == "#N/A"


def test_xl_at_and_out_of_range() -> None:
    assert xl_at((10.0, 20.0, 30.0), 1) == 20.0
    with pytest.raises(XlError) as exc:
        xl_at((10.0,), -1)
    assert exc.value.code == "#VALUE!"
    with pytest.raises(XlError) as exc:
        xl_at((10.0,), 1)
    assert exc.value.code == "#VALUE!"


def test_xl_raise() -> None:
    with pytest.raises(XlError) as exc:
        xl_raise("#N/A")
    assert exc.value.code == "#N/A"
