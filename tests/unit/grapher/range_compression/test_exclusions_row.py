"""Row-axis mirrors of TACO compression exclusion tests."""

from __future__ import annotations

from pathlib import Path

import fastpyxl.utils.cell
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef as GuardCellRef
from excel_grapher.grapher.guard import Compare, Literal
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression import (
    PatternKind,
    RangeRef,
    TacoBuildConfig,
    build_taco_index,
)

from .parity_helpers import assert_taco_parity


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


def test_guarded_edge_row_stays_single() -> None:
    graph = DependencyGraph()
    for col_i in range(6, 9):
        dep_col = fastpyxl.utils.cell.get_column_letter(col_i)
        prec_col = fastpyxl.utils.cell.get_column_letter(col_i - 1)
        graph.add_node(_make_node(f"Sheet1!{prec_col}9", formula=None, is_leaf=True))
        graph.add_node(_make_node(f"Sheet1!{dep_col}9", formula=f"=IF(A1,{prec_col}9,0)"))
        graph.add_edge(
            f"Sheet1!{dep_col}9",
            f"Sheet1!{prec_col}9",
            guard=Compare(GuardCellRef("Sheet1!A1"), ">", Literal(0)),
        )

    index = build_taco_index(graph)
    assert index.compressed_edges == []
    assert len(index.single_edges) > 0
    assert_taco_parity(graph, index)


def test_non_autofill_col_span_not_one_row_group(tmp_path: Path) -> None:
    """A column formula referencing multiple rows must not row-compress as autofill."""
    path = tmp_path / "tall_col.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Demo")
    for row in range(3, 8):
        ws.write_number(row - 1, 3, 1.0)
    ws.write_formula(8, 8, "=D3+D7")
    wb.close()

    graph = create_dependency_graph(path, ["Demo!I9"], load_values=False)
    index = build_taco_index(graph)
    assert not any(
        e.meta.kind in (PatternKind.rr, PatternKind.rr_chain)
        and e.dependent.min_row == e.dependent.max_row
        for e in index.compressed_edges
    )
    assert_taco_parity(graph, index)


def test_static_range_one_off_sum_row_not_compressed() -> None:
    graph = DependencyGraph()
    for col_i in range(5, 8):
        col = fastpyxl.utils.cell.get_column_letter(col_i)
        graph.add_node(_make_node(f"Sheet1!{col}9", formula=None, is_leaf=True))
    graph.add_node(_make_node("Sheet1!H9", formula="=SUM(E9:G9)"))
    dr = DependencyCause.static_range
    for col_i in range(5, 8):
        col = fastpyxl.utils.cell.get_column_letter(col_i)
        graph.add_edge(
            "Sheet1!H9",
            f"Sheet1!{col}9",
            provenance=EdgeProvenance(causes=frozenset({dr})),
        )

    index = build_taco_index(graph)
    assert index.compressed_edges == []
    assert_taco_parity(graph, index)


def test_static_range_breaks_row_rr_group() -> None:
    """A horizontal run whose last cell uses a static range must not compress as one row unit."""
    graph = DependencyGraph()
    dr = DependencyCause.direct_ref
    sr = DependencyCause.static_range
    for col_i in range(4, 9):
        col = fastpyxl.utils.cell.get_column_letter(col_i)
        graph.add_node(_make_node(f"Sheet1!{col}9", formula=None, is_leaf=True))
    graph.add_node(_make_node("Sheet1!F9", formula="=E9"))
    graph.add_node(_make_node("Sheet1!G9", formula="=F9"))
    graph.add_node(_make_node("Sheet1!H9", formula="=SUM(D9:H9)"))
    graph.add_edge(
        "Sheet1!F9",
        "Sheet1!E9",
        provenance=EdgeProvenance(causes=frozenset({dr})),
    )
    graph.add_edge(
        "Sheet1!G9",
        "Sheet1!F9",
        provenance=EdgeProvenance(causes=frozenset({dr})),
    )
    for col_i in range(4, 9):
        col = fastpyxl.utils.cell.get_column_letter(col_i)
        graph.add_edge(
            "Sheet1!H9",
            f"Sheet1!{col}9",
            provenance=EdgeProvenance(causes=frozenset({sr})),
        )

    index = build_taco_index(graph)
    row_edges = [
        e
        for e in index.compressed_edges
        if e.dependent.min_row == e.dependent.max_row == 9
        and e.meta.kind in (PatternKind.rr, PatternKind.rr_chain)
    ]
    assert not any("Sheet1!H9" in e.dependent.cell_keys() for e in row_edges)
    assert_taco_parity(graph, index)


def test_exclude_targets_splits_row_compression() -> None:
    graph = DependencyGraph()
    dr = DependencyCause.direct_ref
    for col_i in range(4, 8):
        prec_col = fastpyxl.utils.cell.get_column_letter(col_i)
        graph.add_node(_make_node(f"Sheet1!{prec_col}9", formula=None, is_leaf=True))
    for col_i in range(6, 9):
        dep_col = fastpyxl.utils.cell.get_column_letter(col_i)
        prec_col = fastpyxl.utils.cell.get_column_letter(col_i - 2)
        graph.add_node(
            _make_node(
                f"Sheet1!{dep_col}9",
                formula=f"={prec_col}9",
                is_target=(dep_col == "H"),
            ),
        )
        graph.add_edge(
            f"Sheet1!{dep_col}9",
            f"Sheet1!{prec_col}9",
            provenance=EdgeProvenance(causes=frozenset({dr})),
        )

    default = build_taco_index(graph)
    rr = [e for e in default.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr) == 1
    assert rr[0].dependent == RangeRef.row_span("Sheet1", 9, "F", "H")

    bounded = build_taco_index(graph, TacoBuildConfig(exclude_targets=True))
    rr_bounded = [e for e in bounded.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr_bounded) == 1
    assert rr_bounded[0].dependent == RangeRef.row_span("Sheet1", 9, "F", "G")
    assert_taco_parity(graph, bounded)
