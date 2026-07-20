"""Stress tests: greedy H-run + V-merge cover stays an exact cell set."""

from __future__ import annotations

from itertools import combinations
from random import Random

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core.address_keys import (
    RangeKey,
    UnionKey,
    expand_node_cells,
    members_to_node_key,
)


def _assert_exact_cover(members: list[str], *, label: str) -> None:
    key = members_to_node_key(members)
    got = {str(c) for c in expand_node_cells(key)}
    want = set(members)
    assert got == want, (
        f"{label}: cover {key!r} is not exact "
        f"(extra={sorted(got - want)}, missing={sorted(want - got)})"
    )


@pytest.mark.parametrize(
    ("label", "members"),
    [
        (
            "L-shape",
            ["Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!A2"],
        ),
        (
            "U-shape",
            [
                "Sheet1!A1",
                "Sheet1!C1",
                "Sheet1!A2",
                "Sheet1!C2",
                "Sheet1!A3",
                "Sheet1!B3",
                "Sheet1!C3",
            ],
        ),
        (
            "row-stripe-with-gap",
            [f"Sheet1!{c}1" for c in list("ABCD") + list("XYZ")],
        ),
        (
            "column-stripe-with-gap",
            [f"Sheet1!A{r}" for r in list(range(1, 6)) + list(range(10, 15))],
        ),
        (
            "missing-corner",
            [
                f"Sheet1!{get_column_letter(c)}{r}"
                for r in range(1, 4)
                for c in range(1, 4)
                if not (c == 3 and r == 3)
            ],
        ),
        (
            "staircase",
            [f"Sheet1!{get_column_letter(i)}{i}" for i in range(1, 8)],
        ),
        (
            "unequal-width-rows",
            [f"Sheet1!{c}1" for c in "ABC"]
            + [f"Sheet1!{c}2" for c in "ABCD"]
            + [f"Sheet1!{c}3" for c in "ABC"],
        ),
        (
            "swiss-cheese",
            [
                f"Sheet1!{get_column_letter(c)}{r}"
                for r in range(1, 5)
                for c in range(1, 5)
                if not (r in (2, 3) and c in (2, 3))
            ],
        ),
        (
            "checkerboard",
            [
                f"Sheet1!{get_column_letter(c)}{r}"
                for r in range(1, 4)
                for c in range(1, 4)
                if (r + c) % 2 == 0
            ],
        ),
    ],
)
def test_adversarial_shapes_are_exact_covers(label: str, members: list[str]) -> None:
    _assert_exact_cover(members, label=label)


def test_unequal_widths_do_not_emit_overcovering_rectangle() -> None:
    members = (
        [f"Sheet1!{c}1" for c in "ABC"]
        + [f"Sheet1!{c}2" for c in "ABCD"]
        + [f"Sheet1!{c}3" for c in "ABC"]
    )
    key = members_to_node_key(members)
    assert isinstance(key, UnionKey)
    # Bounding-box A1:D3 would wrongly include D1 and D3.
    expanded = {str(c) for c in expand_node_cells(key)}
    assert "Sheet1!D1" not in expanded
    assert "Sheet1!D3" not in expanded
    assert "Sheet1!D2" in expanded


def test_contiguous_row_and_column_stripes_collapse_to_single_range() -> None:
    row = [f"Sheet1!{get_column_letter(c)}63" for c in range(4, 26)]  # D..Y
    col = [f"Sheet1!E{r}" for r in range(4, 19)]
    block = [f"Sheet1!{get_column_letter(c)}{r}" for r in range(4, 19) for c in range(5, 10)]

    row_key = members_to_node_key(row)
    col_key = members_to_node_key(col)
    block_key = members_to_node_key(block)

    assert isinstance(row_key, RangeKey) and row_key == "Sheet1!D63:Y63"
    assert isinstance(col_key, RangeKey) and col_key == "Sheet1!E4:E18"
    assert isinstance(block_key, RangeKey) and block_key == "Sheet1!E4:I18"
    for members, key in ((row, row_key), (col, col_key), (block, block_key)):
        assert {str(c) for c in expand_node_cells(key)} == set(members)


def test_exhaustive_2x3_subsets_exact() -> None:
    grid = [f"Sheet1!{get_column_letter(c)}{r}" for r in range(1, 3) for c in range(1, 4)]
    for r in range(1, len(grid) + 1):
        for subset in combinations(grid, r):
            _assert_exact_cover(list(subset), label=f"exh-{subset}")


def test_random_5x5_subsets_exact() -> None:
    rng = Random(0)
    grid = [f"Sheet1!{get_column_letter(c)}{r}" for r in range(1, 6) for c in range(1, 6)]
    for i in range(200):
        k = rng.randint(1, len(grid))
        subset = rng.sample(grid, k)
        _assert_exact_cover(subset, label=f"random-{i}")
