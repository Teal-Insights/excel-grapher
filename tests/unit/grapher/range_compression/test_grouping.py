"""Tests for TACO column- and row-adjacent grouping."""

from __future__ import annotations

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression.config import TacoBuildConfig
from excel_grapher.grapher.range_compression.grouping import (
    adjacent_groups,
    column_adjacent_groups,
    row_adjacent_groups,
)


def _make_node(
    key: str,
    formula: str | None,
    *,
    is_leaf: bool = False,
    is_target: bool = False,
) -> Node:
    sheet, rest = key.split("!", 1)
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=None,
        is_leaf=is_leaf,
        is_target=is_target,
    )


def _group_keys(groups: list[list[str]]) -> set[frozenset[str]]:
    return {frozenset(g) for g in groups}


def test_column_adjacent_groups_vertical_run() -> None:
    graph = DependencyGraph()
    for row in range(3, 8):
        graph.add_node(_make_node(f"Sheet1!D{row}", formula=f"=B{row}"))

    groups = column_adjacent_groups(graph)
    assert groups == [["Sheet1!D3", "Sheet1!D4", "Sheet1!D5", "Sheet1!D6", "Sheet1!D7"]]


def test_column_adjacent_groups_breaks_on_row_gap() -> None:
    graph = DependencyGraph()
    for row in (3, 4, 6, 7):
        graph.add_node(_make_node(f"Sheet1!D{row}", formula=f"=B{row}"))

    groups = column_adjacent_groups(graph)
    assert _group_keys(groups) == {
        frozenset({"Sheet1!D3", "Sheet1!D4"}),
        frozenset({"Sheet1!D6", "Sheet1!D7"}),
    }


def test_row_adjacent_groups_horizontal_run() -> None:
    graph = DependencyGraph()
    for col in ("B", "C", "D", "E", "F"):
        graph.add_node(_make_node(f"Sheet1!{col}9", formula=f"={col}8"))

    groups = row_adjacent_groups(graph)
    assert groups == [
        [
            "Sheet1!B9",
            "Sheet1!C9",
            "Sheet1!D9",
            "Sheet1!E9",
            "Sheet1!F9",
        ]
    ]


def test_row_adjacent_groups_breaks_on_column_gap() -> None:
    graph = DependencyGraph()
    for col in ("B", "C", "E", "F"):
        graph.add_node(_make_node(f"Sheet1!{col}9", formula=f"={col}8"))

    groups = row_adjacent_groups(graph)
    assert _group_keys(groups) == {
        frozenset({"Sheet1!B9", "Sheet1!C9"}),
        frozenset({"Sheet1!E9", "Sheet1!F9"}),
    }


def test_adjacent_groups_column_first_avoids_double_cover() -> None:
    """Cells in a column run are not also claimed by a row run."""
    graph = DependencyGraph()
    for row in range(3, 6):
        graph.add_node(_make_node(f"Sheet1!B{row}", formula=f"=A{row}"))
    for col in ("B", "C", "D"):
        graph.add_node(_make_node(f"Sheet1!{col}3", formula=f"={col}2"))

    groups = adjacent_groups(graph)
    assert _group_keys(groups) == {
        frozenset({"Sheet1!B3", "Sheet1!B4", "Sheet1!B5"}),
        frozenset({"Sheet1!C3", "Sheet1!D3"}),
    }


def test_adjacent_groups_row_first_avoids_double_cover() -> None:
    graph = DependencyGraph()
    for row in range(3, 6):
        graph.add_node(_make_node(f"Sheet1!B{row}", formula=f"=A{row}"))
    for col in ("B", "C", "D"):
        graph.add_node(_make_node(f"Sheet1!{col}3", formula=f"={col}2"))

    groups = adjacent_groups(graph, column_first=False)
    assert _group_keys(groups) == {
        frozenset({"Sheet1!B3", "Sheet1!C3", "Sheet1!D3"}),
        frozenset({"Sheet1!B4", "Sheet1!B5"}),
    }


def test_row_adjacent_groups_exclude_targets_splits_run() -> None:
    graph = DependencyGraph()
    for col in ("B", "C", "D", "E", "F"):
        graph.add_node(
            _make_node(
                f"Sheet1!{col}9",
                formula=f"={col}8",
                is_target=(col == "E"),
            ),
        )

    default = row_adjacent_groups(graph)
    assert default == [
        [
            "Sheet1!B9",
            "Sheet1!C9",
            "Sheet1!D9",
            "Sheet1!E9",
            "Sheet1!F9",
        ]
    ]

    bounded = row_adjacent_groups(
        graph,
        config=TacoBuildConfig(exclude_targets=True),
    )
    assert bounded == [["Sheet1!B9", "Sheet1!C9", "Sheet1!D9"]]
    assert all(len(g) >= 2 for g in bounded)
