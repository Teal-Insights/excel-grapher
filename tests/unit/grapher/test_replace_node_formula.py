"""Topology-aware `replace_node_formula` (#121).

Projection `set_node_formula` still does not rewire. This API is the separate
durable-edit path that re-extracts edges, guards, and provenance.
"""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher.grapher import WorkbookContextRequiredError
from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.graph_consistency import GraphConsistencyKind
from excel_grapher.grapher.guard import CellRef as GuardCellRef
from excel_grapher.grapher.guard import Compare, Literal, Not
from excel_grapher.grapher.node import Node, make_cell_node


def _cell(
    key: str,
    formula: str | None = None,
    *,
    is_leaf: bool | None = None,
    is_target: bool = False,
    value: object = None,
) -> Node:
    sheet, rest = key.split("!", 1)
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    if is_leaf is None:
        is_leaf = formula is None
    return make_cell_node(
        sheet,
        col,
        row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=is_leaf,
        is_target=is_target,
    )


def _direct_edge(graph: DependencyGraph, src: str, dst: str) -> None:
    graph.add_edge(
        src,
        dst,
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )


def test_replace_node_formula_rewires_existing_static_deps() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1", value=1))
    graph.add_node(_cell("Sheet1!C1", value=2))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1", is_target=True))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    graph.replace_node_formula("Sheet1!A1", "=Sheet1!C1", "=Sheet1!C1")

    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!C1"})
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!C1"
    assert node.is_leaf is False
    assert node.is_target is True
    graph.validate_consistency()


def test_set_node_formula_still_does_not_rewire() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    graph.set_node_formula("Sheet1!A1", "=Sheet1!C1", "=Sheet1!C1")

    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})
    issues = graph.consistency_issues()
    kinds = {issue.kind for issue in issues}
    assert GraphConsistencyKind.missing_formula_edge in kinds
    assert GraphConsistencyKind.extra_formula_edge in kinds


def test_replace_clears_formula_and_becomes_leaf() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1", is_target=True))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")
    _direct_edge(graph, "Sheet1!Z1", "Sheet1!A1")

    graph.replace_node_formula("Sheet1!A1", None, None)

    assert graph.get_dependencies("Sheet1!A1") == frozenset()
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.has_formula is False
    assert node.is_leaf is True
    assert graph.get_dependents("Sheet1!A1") == frozenset({"Sheet1!Z1"})


def test_replace_off_path_cell_without_workbook_fails() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    with pytest.raises(WorkbookContextRequiredError, match="missing"):
        graph.replace_node_formula("Sheet1!A1", "=Sheet1!Z9", "=Sheet1!Z9")

    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!B1"


def test_replace_off_path_cell_with_workbook_materializes(tmp_path: Path) -> None:
    path = tmp_path / "off_path.xlsx"
    wb = xlsxwriter.Workbook(path)
    sheet = wb.add_worksheet("Sheet1")
    sheet.write(0, 0, 1)  # A1
    sheet.write(0, 1, 2)  # B1 off-path
    sheet.write_formula(0, 2, "=A1")  # C1
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    assert "Sheet1!B1" not in graph

    graph.replace_node_formula(
        "Sheet1!C1",
        "=B1",
        "=Sheet1!B1",
        workbook=path,
    )

    assert "Sheet1!B1" in graph
    assert graph.get_dependencies("Sheet1!C1") == frozenset({"Sheet1!B1"})
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    assert node.is_leaf is True
    graph.validate_consistency()


def test_replace_named_range_without_workbook_fails(tmp_path: Path) -> None:
    path = tmp_path / "named.xlsx"
    wb = xlsxwriter.Workbook(path)
    sheet = wb.add_worksheet("Sheet1")
    sheet.write(0, 0, 1)  # A1
    sheet.write(0, 1, 99)  # B1 = Rate target, off-path
    sheet.write_formula(0, 2, "=A1")  # C1
    wb.define_name("Rate", "Sheet1!$B$1")
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    assert "Sheet1!B1" not in graph
    assert graph.named_ranges is not None
    assert "Rate" in graph.named_ranges

    with pytest.raises(WorkbookContextRequiredError, match="missing"):
        graph.replace_node_formula("Sheet1!C1", "=Rate", "=Rate")


def test_replace_named_range_unknown_without_maps_fails() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!A1", "=1", is_leaf=True))

    with pytest.raises(WorkbookContextRequiredError, match="named range|defined name"):
        graph.replace_node_formula("Sheet1!A1", "=TaxRate", "=TaxRate")


def test_replace_named_range_with_workbook_materializes(tmp_path: Path) -> None:
    path = tmp_path / "named_ok.xlsx"
    wb = xlsxwriter.Workbook(path)
    sheet = wb.add_worksheet("Sheet1")
    sheet.write(0, 0, 1)  # A1
    sheet.write(0, 1, 99)  # B1
    sheet.write_formula(0, 2, "=A1")  # C1
    wb.define_name("Rate", "Sheet1!$B$1")
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)

    graph.replace_node_formula(
        "Sheet1!C1",
        "=Rate",
        "=Rate",
        workbook=path,
    )

    assert graph.get_dependencies("Sheet1!C1") == frozenset({"Sheet1!B1"})
    assert "Sheet1!B1" in graph
    graph.validate_consistency()


def test_replace_dynamic_ref_without_workbook_fails(tmp_path: Path) -> None:
    path = tmp_path / "dyn.xlsx"
    wb = xlsxwriter.Workbook(path)
    sheet = wb.add_worksheet("Sheet1")
    sheet.write(0, 0, 10)  # A1
    sheet.write(0, 1, 1)  # B1 offset
    sheet.write(1, 0, 20)  # A2 target
    sheet.write_formula(0, 2, "=A1")  # C1
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)

    with pytest.raises(WorkbookContextRequiredError, match="OFFSET|dynamic"):
        graph.replace_node_formula(
            "Sheet1!C1", "=OFFSET(A1,B1,0)", "=OFFSET(Sheet1!A1,Sheet1!B1,0)"
        )


def test_replace_dynamic_ref_with_workbook(tmp_path: Path) -> None:
    path = tmp_path / "dyn_ok.xlsx"
    wb = xlsxwriter.Workbook(path)
    sheet = wb.add_worksheet("Sheet1")
    sheet.write(0, 0, 10)  # A1
    sheet.write(0, 1, 1)  # B1
    sheet.write(1, 0, 20)  # A2
    sheet.write_formula(0, 2, "=A1", None, 10)  # C1
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=True)

    graph.replace_node_formula(
        "Sheet1!C1",
        "=OFFSET(A1,B1,0)",
        "=OFFSET(Sheet1!A1,Sheet1!B1,0)",
        workbook=path,
        use_cached_dynamic_refs=True,
    )

    deps = graph.get_dependencies("Sheet1!C1")
    assert "Sheet1!A2" in deps
    assert "Sheet1!B1" in deps
    assert "Sheet1!A2" in graph
    attrs = graph.get_edge_attrs("Sheet1!C1", "Sheet1!A2")
    assert attrs.provenance is not None
    assert attrs.provenance.causes & DependencyCause.dynamic_offset


def test_replace_missing_node_raises_key_error() -> None:
    graph = DependencyGraph()
    with pytest.raises(KeyError, match="Sheet1!A1"):
        graph.replace_node_formula("Sheet1!A1", "=1", "=1")


def test_replace_preserves_incoming_edges() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!C1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    graph.add_node(_cell("Sheet1!D1", "=Sheet1!A1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")
    _direct_edge(graph, "Sheet1!D1", "Sheet1!A1")

    graph.replace_node_formula("Sheet1!A1", "=Sheet1!C1", "=Sheet1!C1")

    assert graph.get_dependents("Sheet1!A1") == frozenset({"Sheet1!D1"})
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!C1"})


def test_replace_if_formula_rewires_guards_with_workbook(tmp_path: Path) -> None:
    path = tmp_path / "iff.xlsx"
    wb = xlsxwriter.Workbook(path)
    sheet = wb.add_worksheet("Sheet1")
    sheet.write(0, 0, 1)  # A1
    sheet.write(0, 1, 1)  # B1 cond
    sheet.write(0, 2, 2)  # C1 then
    sheet.write(0, 3, 3)  # D1 else
    sheet.write_formula(0, 4, "=A1")  # E1
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!E1"], load_values=False)

    graph.replace_node_formula(
        "Sheet1!E1",
        "=IF(B1>0,C1,D1)",
        "=IF(Sheet1!B1>0,Sheet1!C1,Sheet1!D1)",
        workbook=path,
    )

    assert graph.get_dependencies("Sheet1!E1") == frozenset({"Sheet1!B1", "Sheet1!C1", "Sheet1!D1"})
    then_guard = graph.get_edge_guard("Sheet1!E1", "Sheet1!C1")
    else_guard = graph.get_edge_guard("Sheet1!E1", "Sheet1!D1")
    cond = Compare(GuardCellRef("Sheet1!B1"), ">", Literal(0))
    assert then_guard == cond
    assert else_guard == Not(cond)
    assert graph.get_edge_guard("Sheet1!E1", "Sheet1!B1") is None
    graph.validate_consistency()


def test_replace_off_path_formula_cell_extracts_its_subgraph(tmp_path: Path) -> None:
    path = tmp_path / "nested.xlsx"
    wb = xlsxwriter.Workbook(path)
    sheet = wb.add_worksheet("Sheet1")
    sheet.write(0, 0, 1)  # A1
    sheet.write(0, 1, 2)  # B1
    sheet.write_formula(0, 2, "=B1*2")  # C1 off-path formula
    sheet.write_formula(0, 3, "=A1")  # D1 target
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)
    assert "Sheet1!C1" not in graph
    assert "Sheet1!B1" not in graph

    graph.replace_node_formula(
        "Sheet1!D1",
        "=C1",
        "=Sheet1!C1",
        workbook=path,
    )

    assert "Sheet1!C1" in graph
    assert "Sheet1!B1" in graph
    assert graph.get_dependencies("Sheet1!D1") == frozenset({"Sheet1!C1"})
    assert graph.get_dependencies("Sheet1!C1") == frozenset({"Sheet1!B1"})
    graph.validate_consistency()
