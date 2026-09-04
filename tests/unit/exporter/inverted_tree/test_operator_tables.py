"""Rung-3 operators are the evaluator's `core.operators` table (#656, #651)."""

from __future__ import annotations

from collections.abc import Callable
from typing import cast

from excel_grapher.core.operators import OPERATOR_TABLE as CORE_OPERATOR_TABLE
from excel_grapher.core.types import XlError as CoreXlError
from excel_grapher.exporter.inverted_tree import runtime as inverted_runtime

_Binary = Callable[[object, object], object]


def test_operator_tables_have_the_same_keys() -> None:
    assert set(inverted_runtime.OPERATOR_TABLE) == set(CORE_OPERATOR_TABLE)


def test_operator_tables_agree_on_scalar_samples() -> None:
    samples = [
        ("+", 2, 3, 5.0),
        ("-", 5, 2, 3.0),
        ("*", 4, 2.5, 10.0),
        ("/", 9, 3, 3.0),
        ("^", 2, 3, 8.0),
        ("=", "Nominal", "nominal", True),
        ("=", True, 1, False),
        (">", True, 100, True),
        ("=", "10", 10, False),
        ("=", "", 0, False),
        ("<", "a", 1, False),
        ("<>", "abc", 1, True),
    ]
    for op, left, right, expected in samples:
        core = cast(_Binary, CORE_OPERATOR_TABLE[op])(left, right)
        if isinstance(core, CoreXlError):
            raise AssertionError(f"core {op} returned {core}")
        inverted = cast(_Binary, inverted_runtime.OPERATOR_TABLE[op])(left, right)
        assert inverted == core == expected, (op, left, right, inverted, core)
