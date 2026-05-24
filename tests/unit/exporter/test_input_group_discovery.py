"""Tests for CodeGenerator.discover_input_groups."""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.input_group_discovery import validate_input_groups
from excel_grapher.exporter.input_groups import GroupingOptions, GroupingOverride


def _make_node(
    address: str,
    formula: str | None,
    value: object,
    *,
    metadata: dict | None = None,
) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
        metadata=dict(metadata or {}),
    )


def _graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_discover_input_groups_singleton_without_labels() -> None:
    graph = _graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=S!A1*2", None),
    )
    payload = CodeGenerator(graph).discover_input_groups(["S!B1"])
    assert payload.summary.total_groups == 1
    assert payload.summary.total_cells == 1
    assert payload.groups[0].cells[0].address == "S!A1"
    validate_input_groups(payload.groups)


def test_discover_input_groups_stable_group_id() -> None:
    graph = _graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=S!A1*2", None),
    )
    gen = CodeGenerator(graph)
    p1 = gen.discover_input_groups(["S!B1"])
    p2 = gen.discover_input_groups(["S!B1"])
    assert p1.groups[0].group_id == p2.groups[0].group_id


def test_discover_input_groups_rectangular_metadata() -> None:
    graph = _graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", None, 2.0),
        _make_node("S!A2", None, 3.0),
        _make_node("S!B2", None, 4.0),
        _make_node("S!C1", "=S!A1+S!B1+S!A2+S!B2", None),
    )
    payload = CodeGenerator(graph).discover_input_groups(
        ["S!C1"],
        grouping=GroupingOptions(
            include_labels=True,
            label_mode="first",
            overrides=(
                GroupingOverride(
                    range_spec="S!A1:B2",
                    orientation="columnwise",
                ),
            ),
        ),
    )
    assert payload.summary.total_groups == 1
    group = payload.groups[0]
    assert group.orientation == "columnwise"
    assert group.shape == (2, 2)
    assert group.range_a1 == "S!A1:S!B2"
    assert group.bounding_box is not None


def test_discover_groups_by_shared_row_labels() -> None:
    graph = _graph(
        _make_node(
            "S!B1",
            None,
            1.0,
            metadata={"row_labels": ["Revenue"], "column_labels": ["2021"]},
        ),
        _make_node(
            "S!C1",
            None,
            2.0,
            metadata={"row_labels": ["Revenue"], "column_labels": ["2022"]},
        ),
        _make_node("S!D1", "=S!B1+S!C1", None),
    )
    payload = CodeGenerator(graph).discover_input_groups(
        ["S!D1"],
        grouping=GroupingOptions(include_labels=True, label_mode="all"),
    )
    assert payload.summary.total_groups == 1
    assert len(payload.groups[0].cells) == 2


def test_override_changes_orientation_on_rediscover() -> None:
    graph = _graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=S!A1*2", None),
    )
    gen = CodeGenerator(graph)
    neutral = gen.discover_input_groups(["S!B1"])
    columnwise = gen.discover_input_groups(
        ["S!B1"],
        grouping=GroupingOptions(
            overrides=(GroupingOverride(range_spec="S!A1", orientation="columnwise"),),
        ),
    )
    assert neutral.groups[0].orientation == "rowwise"
    assert columnwise.groups[0].orientation == "columnwise"
