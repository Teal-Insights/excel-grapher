"""TACO compression boundaries for codegen (targets, inputs, internal-only)."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression import (
    PatternKind,
    TacoBuildConfig,
    build_taco_index,
    input_keys_from_graph,
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


def test_default_config_unchanged_rr_compression() -> None:
    graph = DependencyGraph()
    for row in range(3, 8):
        graph.add_node(_make_node(f"Sheet1!B{row}", formula=None, is_leaf=True))
        graph.add_node(_make_node(f"Sheet1!D{row}", formula=f"=B{row}"))
        graph.add_edge(f"Sheet1!D{row}", f"Sheet1!B{row}")
    index = build_taco_index(graph)
    assert any(e.meta.kind == PatternKind.rr for e in index.compressed_edges)


def test_exclude_targets_splits_column_group() -> None:
    graph = DependencyGraph()
    for row in range(3, 8):
        graph.add_node(_make_node(f"Sheet1!B{row}", formula=None, is_leaf=True))
        graph.add_node(
            _make_node(
                f"Sheet1!D{row}",
                formula=f"=B{row}",
                is_target=(row == 7),
            ),
        )
        graph.add_edge(f"Sheet1!D{row}", f"Sheet1!B{row}")

    default = build_taco_index(graph)
    assert len([e for e in default.compressed_edges if e.meta.kind == PatternKind.rr]) == 1

    bounded = build_taco_index(
        graph,
        TacoBuildConfig(exclude_targets=True),
    )
    rr = [e for e in bounded.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr) == 1
    assert rr[0].dependent.max_row == 6
    assert_taco_parity(graph, bounded)


def test_exclude_input_keys_skips_precedent_compression() -> None:
    graph = DependencyGraph()
    input_keys: set[str] = set()
    for row in range(3, 8):
        b_key = f"Sheet1!B{row}"
        graph.add_node(_make_node(b_key, formula=None, is_leaf=True))
        input_keys.add(b_key)
        graph.add_node(_make_node(f"Sheet1!C{row}", formula=None, is_leaf=True))
        graph.add_node(_make_node(f"Sheet1!D{row}", formula=f"=B{row}*C{row}"))
        graph.add_edge(f"Sheet1!D{row}", b_key)
        graph.add_edge(f"Sheet1!D{row}", f"Sheet1!C{row}")

    config = TacoBuildConfig(exclude_input_keys=frozenset(input_keys))
    index = build_taco_index(graph, config)
    rr = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr) == 1
    assert rr[0].precedent.min_col == "C"
    assert_taco_parity(graph, index)


def test_internal_only_compresses_middle_formula_column(tmp_path: Path) -> None:
    """Inputs and targets stay single; only internal formula columns compress."""
    path = tmp_path / "internal_chain.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Model")
    for row in range(3, 8):
        ws.write_number(row - 1, 0, float(row))  # A inputs
        ws.write_formula(row - 1, 1, f"=A{row}*2")  # B internal
        ws.write_formula(row - 1, 2, f"=B{row}+1")  # C internal
        ws.write_formula(row - 1, 3, f"=C{row}")  # D targets
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Model!D3:D7"],
        load_values=False,
        store_raw_formula=True,
    )
    inputs = frozenset(f"Model!A{row}" for row in range(3, 8))
    index = build_taco_index(
        graph,
        TacoBuildConfig.for_codegen(input_keys=inputs),
    )
    rr = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr) == 1
    assert rr[0].dependent.min_col == "C"
    assert rr[0].precedent.min_col == "B"
    for edge in index.compressed_edges:
        assert edge.dependent.min_col not in {"A", "D"}
        assert edge.precedent.min_col != "A"
    assert_taco_parity(graph, index)


def test_input_keys_from_graph_leaf_classification() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("Sheet1!A1", formula=None, is_leaf=True))
    graph.leaf_classification = {"Sheet1!A1": "input", "Sheet1!B1": "constant"}
    assert input_keys_from_graph(graph) == frozenset({"Sheet1!A1"})
