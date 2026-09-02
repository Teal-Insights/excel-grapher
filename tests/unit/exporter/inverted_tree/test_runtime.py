"""Runtime primitives for inverted-tree codegen (`take`, `require_aligned`, `xl_*`)."""

from __future__ import annotations

import pytest

from excel_grapher.exporter.inverted_tree.runtime import (
    InstanceCycleError,
    XlError,
    as_measure,
    demand_instance,
    eval_instance,
    is_error,
    live_measure,
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


def test_take_accepts_range() -> None:
    values = (10, 20, 30, 40, 50)
    assert take(values, range(0, 3)) == (10, 20, 30)
    assert take(values, range(1, 5, 2)) == (20, 40)


def test_take_accepts_slice() -> None:
    values = (10, 20, 30, 40, 50)
    assert take(values, slice(0, 3)) == (10, 20, 30)
    assert take(values, slice(1, 5, 2)) == (20, 40)
    assert take(values, slice(None, 2)) == (10, 20)
    assert take(values, slice(3, None)) == (40, 50)


def test_take_slice_fails_closed_when_stop_exceeds_length() -> None:
    with pytest.raises(ValueError, match="take index 3"):
        take((1, 2, 3), slice(0, 4))


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


def test_as_measure_preserves_error_codes() -> None:
    assert as_measure(1.5) == 1.5
    assert as_measure("#REF!") == "#REF!"
    assert as_measure("#DIV/0!") == "#DIV/0!"
    assert is_error("#REF!")
    assert not is_error(1.5)
    assert not is_error("REF")


def test_eval_instance_memos_and_detects_cycles() -> None:
    memo: dict[tuple[str, int], float] = {}
    stack: set[tuple[str, int]] = set()
    calls = {"n": 0}

    def compute(index: int) -> float:
        calls["n"] += 1
        if index == 0:
            return 1.0
        return eval_instance("s", index - 1, compute, memo, stack) + 1.0

    assert eval_instance("s", 2, compute, memo, stack) == 3.0
    assert eval_instance("s", 2, compute, memo, stack) == 3.0
    assert calls["n"] == 3

    def loop(_index: int) -> float:
        return eval_instance("c", 0, loop, memo, stack)

    with pytest.raises(InstanceCycleError, match="distance-zero cycle"):
        eval_instance("c", 0, loop, memo, stack)


def test_demand_instance_reraises_stored_error_codes() -> None:
    memo: dict[tuple[str, int], float | str] = {}
    stack: set[tuple[str, int]] = set()

    def compute(index: int) -> float | str:
        del index
        return "#REF!"

    with pytest.raises(XlError) as exc:
        demand_instance("s", 0, compute, memo, stack)
    assert exc.value.code == "#REF!"
    assert eval_instance("s", 0, compute, memo, stack) == "#REF!"


def test_live_measure_reraises_error_codes() -> None:
    assert live_measure(1.5) == 1.5
    with pytest.raises(XlError) as exc:
        live_measure("#REF!")
    assert exc.value.code == "#REF!"
