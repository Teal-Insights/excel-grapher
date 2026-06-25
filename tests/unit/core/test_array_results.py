"""Tests for top-level array result helpers."""

from __future__ import annotations

import numpy as np

from excel_grapher.core.array_results import (
    array_values_equal,
    finalize_top_level_array_result,
    is_array_result,
    read_spill_scalar,
    spill_blocked_at_anchor,
    spill_footprint_addresses,
    spill_offsets,
)
from excel_grapher.core.types import XlError


def test_is_array_result_true_for_multicell_ndarray() -> None:
    value = np.array([[True], [False]], dtype=object)
    assert is_array_result(value)


def test_is_array_result_false_for_scalar_and_1x1() -> None:
    assert not is_array_result(True)
    assert not is_array_result(np.array([[1.0]], dtype=object))


def test_array_values_equal_matches_bool_columns() -> None:
    left = np.array([[True], [False], [True]], dtype=object)
    right = np.array([[True], [False], [True]], dtype=object)
    assert array_values_equal(left, right)


def test_array_values_equal_rejects_shape_mismatch() -> None:
    left = np.array([[True], [False]], dtype=object)
    right = np.array([[True, False]], dtype=object)
    assert not array_values_equal(left, right)


def test_array_values_equal_compares_xlerror_cells() -> None:
    left = np.array([[XlError.VALUE, 1]], dtype=object)
    right = np.array([[XlError.VALUE, 1]], dtype=object)
    assert array_values_equal(left, right)


def test_spill_offsets_column_vector() -> None:
    assert spill_offsets(10, 4, 11, 4, (3, 1)) == (1, 0)
    assert spill_offsets(10, 4, 12, 5, (3, 1)) is None


def test_read_spill_scalar_from_cached_anchor() -> None:
    anchor = np.array([[True], [False], [True]], dtype=object)
    cache = {"Data!D10": anchor}
    assert read_spill_scalar("Data!D11", cache) is False
    assert read_spill_scalar("Data!D12", cache) is True


def test_spill_footprint_addresses_column() -> None:
    assert spill_footprint_addresses("Data!D10", (3, 1)) == [
        "Data!D10",
        "Data!D11",
        "Data!D12",
    ]


def test_spill_footprint_addresses_2d() -> None:
    assert spill_footprint_addresses("Data!D10", (2, 2)) == [
        "Data!D10",
        "Data!E10",
        "Data!D11",
        "Data!E11",
    ]


def test_finalize_top_level_array_result_spill_blocked() -> None:
    array = np.array([[True], [False], [True]], dtype=object)
    result = finalize_top_level_array_result(
        "Data!D10",
        array,
        is_occupied=lambda addr: addr == "Data!D11",
    )
    assert result is XlError.SPILL


def test_spill_blocked_at_anchor_ignores_anchor_itself() -> None:
    assert not spill_blocked_at_anchor(
        "Data!D10",
        (3, 1),
        is_occupied=lambda addr: addr == "Data!D10",
    )
