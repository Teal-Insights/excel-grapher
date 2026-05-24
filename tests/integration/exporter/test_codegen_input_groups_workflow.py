"""End-to-end discover → inspect/edit → generate workflow tests."""

from __future__ import annotations

import sys

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.input_groups import (
    GroupingOptions,
    GroupingOverride,
    InputGroup,
    InputGroupsPayload,
    SetterGenerationOptions,
)
from tests.integration.utils.generated_package import import_generated_package, purge_module_cache


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


def test_discover_serialize_rehydrate_generate_setters() -> None:
    graph = _graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=S!A1*2", None),
    )
    gen = CodeGenerator(graph)
    payload = gen.discover_input_groups(["S!B1"])
    restored = InputGroupsPayload.from_dict(payload.to_dict())
    setters_code = gen.generate_setters_module(restored.groups)
    assert "def set_" in setters_code
    assert "S!A1" in setters_code


def test_explicit_input_groups_bypasses_rediscovery() -> None:
    graph = _graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", None, 2.0),
        _make_node("S!C1", "=S!A1+S!B1", None),
    )
    gen = CodeGenerator(graph)
    auto = gen.discover_input_groups(["S!C1"])
    assert auto.summary.total_groups == 2

    manual_group = InputGroup(
        group_id="manual_inputs",
        sheet="S",
        orientation="rowwise",
        row_labels_key=(),
        column_labels_key=(),
        cells=(
            auto.groups[0].cells[0],
            auto.groups[1].cells[0],
        ),
        bounding_box=None,
        shape=None,
        range_a1=None,
    )
    files = gen.generate_modules(
        ["S!C1"],
        package_name="exported_manual_groups",
        setters=SetterGenerationOptions(),
        input_groups=(manual_group,),
    )
    setters_py = files["exported_manual_groups/setters.py"]
    assert "manual_inputs" in setters_py
    assert "def set_manual_inputs" in setters_py


def test_discover_override_rediscover_changes_orientation() -> None:
    graph = _graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=S!A1*2", None),
    )
    gen = CodeGenerator(graph)
    before = gen.discover_input_groups(["S!B1"])
    after = gen.discover_input_groups(
        ["S!B1"],
        grouping=GroupingOptions(
            overrides=(GroupingOverride(range_spec="S!A1", orientation="columnwise"),),
        ),
    )
    assert before.groups[0].orientation == "rowwise"
    assert after.groups[0].orientation == "columnwise"
    assert before.groups[0].group_id != after.groups[0].group_id


def test_label_mode_first_projects_single_label_field() -> None:
    graph = _graph(
        _make_node(
            "S!A1",
            None,
            1.0,
            metadata={"row_labels": ["Revenue"], "column_labels": ["2021"]},
        ),
        _make_node("S!B1", "=S!A1*2", None),
    )
    payload = CodeGenerator(graph).discover_input_groups(
        ["S!B1"],
        grouping=GroupingOptions(include_labels=True, label_mode="first"),
    )
    cell = payload.groups[0].cells[0]
    assert cell.row_labels == ("Revenue",)
    assert cell.column_labels == ("2021",)


def test_generate_modules_setters_roundtrip(tmp_path) -> None:
    graph = _graph(
        _make_node("S!A1", None, 10.0),
        _make_node("S!B1", "=S!A1*2", None),
    )
    gen = CodeGenerator(graph)
    payload = gen.discover_input_groups(["S!B1"])
    package_name = "exported_workflow_setters"
    files = gen.generate_modules(
        ["S!B1"],
        package_name=package_name,
        setters=SetterGenerationOptions(),
        input_groups=payload.groups,
    )
    pkg = import_generated_package(tmp_path, files, package_name=package_name)
    try:
        setter_name = next(
            name for name in dir(pkg.setters) if name.startswith("set_") and name != "set_inputs"
        )
        ctx = pkg.make_context()
        getattr(pkg.setters, setter_name)(ctx, [{"address": "S!A1", "value": 5.0}])
        records = pkg.compute_all(ctx=ctx)
        values = {rec["address"]: rec["value"] for rec in records if "address" in rec}
        assert values["S!B1"] == 10.0
    finally:
        sys.path.remove(str(tmp_path))
        purge_module_cache(package_name)
