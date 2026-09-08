"""Fail-closed formula/edge consistency checks for projected graphs (#261)."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    CellRef,
    CellRefNode,
    RelativeAxis,
)
from excel_grapher.exporter.projection import (
    IdentityTransitCompression,
    ProjectionResult,
    apply_projection,
)
from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.graph_consistency import (
    GraphConsistencyError,
    GraphConsistencyKind,
)
from excel_grapher.grapher.node import Node, make_cell_node
from excel_grapher.grapher.parser import DEFAULT_MAX_RANGE_CELLS


def _cell(
    key: str,
    formula: str | None = None,
    *,
    is_leaf: bool | None = None,
    is_target: bool = False,
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
        is_leaf=is_leaf,
        is_target=is_target,
    )


def _direct_edge(graph: DependencyGraph, src: str, dst: str) -> None:
    graph.add_edge(
        src,
        dst,
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )


def _issue_tuples(exc: GraphConsistencyError) -> set[tuple[str, str | None, str | None]]:
    return {(issue.kind.value, issue.from_key, issue.to_key) for issue in exc.issues}


def test_consistent_cell_ref_graph_passes() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    graph.validate_consistency()
    assert graph.consistency_issues() == ()


def test_set_node_formula_without_rewire_is_inconsistent() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!C1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    graph.set_node_formula("Sheet1!A1", "=Sheet1!C1", "=Sheet1!C1")

    issues = graph.consistency_issues()
    kinds = {issue.kind for issue in issues}
    assert GraphConsistencyKind.missing_formula_edge in kinds
    assert GraphConsistencyKind.extra_formula_edge in kinds
    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    tuples = _issue_tuples(exc_info.value)
    assert (
        GraphConsistencyKind.missing_formula_edge.value,
        "Sheet1!A1",
        "Sheet1!C1",
    ) in tuples
    assert (
        GraphConsistencyKind.extra_formula_edge.value,
        "Sheet1!A1",
        "Sheet1!B1",
    ) in tuples


def test_set_node_formula_does_not_auto_rewire_or_validate() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    graph.set_node_formula("Sheet1!A1", "=Sheet1!C1", "=Sheet1!C1")

    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!C1"


def test_remove_node_leaves_stale_formula_ref() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    graph.remove_node("Sheet1!B1")

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.missing_formula_edge.value,
        "Sheet1!A1",
        "Sheet1!B1",
    ) in _issue_tuples(exc_info.value)


def test_range_formula_requires_expanded_member_edges() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!A1"))
    graph.add_node(_cell("Sheet1!A2"))
    graph.add_node(_cell("Sheet1!A3"))
    graph.add_node(_cell("Sheet1!B1", "=SUM(Sheet1!A1:A3)"))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A1")
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A3")

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.missing_formula_edge.value,
        "Sheet1!B1",
        "Sheet1!A2",
    ) in _issue_tuples(exc_info.value)


def test_dynamic_only_extra_edge_is_not_formula_ref_disagreement() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!A2"))
    graph.add_node(_cell("Sheet1!A3"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!A2"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!A2")
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!A3",
        provenance=EdgeProvenance(causes=DependencyCause.dynamic_offset),
    )

    graph.validate_consistency()


def test_edge_to_missing_node() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!Z9"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!Z9")

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.edge_to_missing_node.value,
        "Sheet1!A1",
        "Sheet1!Z9",
    ) in _issue_tuples(exc_info.value)


def test_edge_from_missing_node() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.edge_from_missing_node.value,
        "Sheet1!A1",
        "Sheet1!B1",
    ) in _issue_tuples(exc_info.value)


def test_leaf_flag_mismatch_when_outgoing_edges_exist() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1", is_leaf=True))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.leaf_flag_mismatch.value,
        "Sheet1!A1",
        None,
    ) in _issue_tuples(exc_info.value)


def test_leaf_flag_mismatch_when_formula_has_no_deps() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!A1", "=1", is_leaf=False))

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.leaf_flag_mismatch.value,
        "Sheet1!A1",
        None,
    ) in _issue_tuples(exc_info.value)


def test_formula_flag_mismatch_for_value_cell_with_edges() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", is_leaf=False))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.formula_flag_mismatch.value,
        "Sheet1!A1",
        None,
    ) in _issue_tuples(exc_info.value)


def test_extracted_workbook_is_consistent(tmp_path: Path) -> None:
    path = tmp_path / "tiny.xlsx"
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet("Sheet1")
    sheet.write(0, 0, 1)
    sheet.write_formula(0, 1, "=A1+1")
    workbook.close()

    graph = create_dependency_graph(path, ["Sheet1!B1"], load_values=False)
    graph.validate_consistency()


def test_identity_projection_stays_consistent() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!C1"))
    graph.add_node(_cell("Sheet1!B1", "=Sheet1!C1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1", is_target=True))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    graph.validate_consistency()
    projection = IdentityTransitCompression().project(graph)
    projection.validate()


def test_apply_projection_validate_opt_in_raises() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    class StaleFormula:
        def project(self, source: DependencyGraph) -> ProjectionResult:
            projected = source.copy()
            projected.set_node_formula("Sheet1!A1", "=Sheet1!C1", "=Sheet1!C1")
            from excel_grapher.exporter.projection import BaseProjectionManifest

            return ProjectionResult(
                source,
                projected,
                BaseProjectionManifest.empty(kind="stale"),
            )

    result = apply_projection(graph, [StaleFormula()])
    with pytest.raises(GraphConsistencyError):
        result.validate()
    with pytest.raises(GraphConsistencyError):
        apply_projection(graph, [StaleFormula()], validate=True)


def test_apply_projection_default_does_not_validate() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    class StaleFormula:
        def project(self, source: DependencyGraph) -> ProjectionResult:
            projected = source.copy()
            projected.set_node_formula("Sheet1!A1", "=Sheet1!C1", "=Sheet1!C1")
            from excel_grapher.exporter.projection import BaseProjectionManifest

            return ProjectionResult(
                source,
                projected,
                BaseProjectionManifest.empty(kind="stale"),
            )

    result = apply_projection(graph, [StaleFormula()])
    assert result.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})


def test_unlabeled_stale_edge_after_extracted_rewrite(tmp_path: Path) -> None:
    path = tmp_path / "rewrite.xlsx"
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet("Sheet1")
    sheet.write(0, 0, 1)
    sheet.write(0, 1, 2)
    sheet.write_formula(0, 2, "=A1")
    workbook.close()

    graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    assert graph.get_edge_attrs("Sheet1!C1", "Sheet1!A1").provenance is None
    graph.set_node_formula("Sheet1!C1", "=Sheet1!B1", "=Sheet1!B1")

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    tuples = _issue_tuples(exc_info.value)
    assert (
        GraphConsistencyKind.missing_formula_edge.value,
        "Sheet1!C1",
        "Sheet1!B1",
    ) in tuples
    assert (
        GraphConsistencyKind.extra_formula_edge.value,
        "Sheet1!C1",
        "Sheet1!A1",
    ) in tuples


def test_offset_extracted_graph_is_consistent(tmp_path: Path) -> None:
    path = tmp_path / "offset.xlsx"
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet("Sheet1")
    sheet.write(0, 0, 1)
    sheet.write(1, 0, 2)
    sheet.write(2, 0, 3)
    sheet.write_formula(0, 1, "=OFFSET(A1:A3,1,0)")
    workbook.close()

    graph = create_dependency_graph(
        path, ["Sheet1!B1"], load_values=False, use_cached_dynamic_refs=True
    )
    graph.validate_consistency()


def test_unlabeled_extra_edge_on_static_formula() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1"))
    graph.add_node(_cell("Sheet1!C1"))
    graph.add_node(_cell("Sheet1!A1", "=Sheet1!B1"))
    graph.add_edge("Sheet1!A1", "Sheet1!B1")
    graph.add_edge("Sheet1!A1", "Sheet1!C1")

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.extra_formula_edge.value,
        "Sheet1!A1",
        "Sheet1!C1",
    ) in _issue_tuples(exc_info.value)


def test_range_too_large_is_issue() -> None:
    graph = DependencyGraph()
    end = DEFAULT_MAX_RANGE_CELLS + 1
    graph.add_node(_cell("Sheet1!B1", f"=SUM(Sheet1!A1:A{end})", is_leaf=True))

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.range_too_large.value,
        "Sheet1!B1",
        None,
    ) in _issue_tuples(exc_info.value)


def test_unbounded_whole_column_without_sheet_bounds() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!B1", "=SUM(C:C)", is_leaf=True))

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.unbounded_whole_ref.value,
        "Sheet1!B1",
        None,
    ) in _issue_tuples(exc_info.value)


def test_unresolved_relative_ref_is_issue() -> None:
    graph = DependencyGraph()
    graph.add_node(_cell("Sheet1!A1", is_leaf=True))
    graph.set_node_ast(
        "Sheet1!A1",
        CellRefNode(CellRef(sheet="Sheet1", col=AbsoluteAxis(1), row=RelativeAxis(-5))),
        formula="=A[-5]",
    )

    with pytest.raises(GraphConsistencyError) as exc_info:
        graph.validate_consistency()
    assert (
        GraphConsistencyKind.unresolved_formula_ref.value,
        "Sheet1!A1",
        None,
    ) in _issue_tuples(exc_info.value)
