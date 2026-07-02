"""TACO compression boundaries for codegen (targets, inputs, internal-only)."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import cast

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression import (
    PatternKind,
    TacoBuildConfig,
    assert_codegen_index_boundaries,
    build_codegen_taco_index,
    build_taco_index,
    codegen_boundary_keys,
    input_keys_from_graph,
    input_keys_from_ranges,
    setter_keys_from_bindings,
)
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from tests.integration.user_flows.test_column_bindings import BINDINGS_DOCUMENT
from tests.integration.user_flows.utils import write_series_bindings_workbook

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


def test_input_keys_from_ranges_expands_rectangles() -> None:
    keys = input_keys_from_ranges(["Sheet1!B3:C4", "Sheet1!E5"])
    assert keys == frozenset(
        {
            "Sheet1!B3",
            "Sheet1!B4",
            "Sheet1!C3",
            "Sheet1!C4",
            "Sheet1!E5",
        }
    )


def test_codegen_boundary_keys_unions_graph_ranges_and_setters(tmp_path: Path) -> None:
    workbook = tmp_path / "bindings.xlsx"
    write_series_bindings_workbook(workbook)
    graph = create_dependency_graph(
        workbook,
        ["Sheet1!F3:F5"],
        load_values=True,
    )
    graph.leaf_classification = {"Sheet1!A1": "input"}
    bindings = cast(WorkbookSeriesBindings, deepcopy(BINDINGS_DOCUMENT))
    bindings["workbook"] = workbook.name

    keys = codegen_boundary_keys(
        graph,
        input_ranges=["Sheet1!G3:G5"],
        series_bindings=bindings,
        bindings_workbook=workbook,
    )
    assert "Sheet1!A1" in keys
    assert "Sheet1!G3" in keys
    setter_keys = setter_keys_from_bindings(graph, bindings, workbook)
    assert "Sheet1!F3" in setter_keys
    assert setter_keys.issubset(keys)


def test_for_codegen_export_asserts_boundaries_on_internal_chain(tmp_path: Path) -> None:
    path = tmp_path / "internal_chain.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Model")
    for row in range(3, 8):
        ws.write_number(row - 1, 0, float(row))
        ws.write_formula(row - 1, 1, f"=A{row}*2")
        ws.write_formula(row - 1, 2, f"=B{row}+1")
        ws.write_formula(row - 1, 3, f"=C{row}")
    wb.close()

    graph = create_dependency_graph(path, ["Model!D3:D7"], load_values=False)
    config = TacoBuildConfig.for_codegen_export(
        graph,
        input_ranges=[f"Model!A{row}" for row in range(3, 8)],
    )
    index = build_taco_index(graph, config)
    assert_codegen_index_boundaries(graph, index, config)
    assert_taco_parity(graph, index)


def test_build_codegen_taco_index_attaches_to_graph(tmp_path: Path) -> None:
    path = tmp_path / "internal_chain.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Model")
    for row in range(3, 8):
        ws.write_number(row - 1, 0, float(row))
        ws.write_formula(row - 1, 1, f"=A{row}*2")
        ws.write_formula(row - 1, 3, f"=B{row}")
    wb.close()

    graph = create_dependency_graph(path, ["Model!D3:D7"], load_values=False)
    index = build_codegen_taco_index(
        graph,
        input_ranges=[f"Model!A{row}" for row in range(3, 8)],
        attach_to_graph=True,
    )
    assert graph.codegen_taco_index is index
    assert_codegen_index_boundaries(
        graph,
        index,
        TacoBuildConfig.for_codegen_export(
            graph,
            input_ranges=[f"Model!A{row}" for row in range(3, 8)],
        ),
    )


def test_code_generator_build_codegen_taco_index(tmp_path: Path) -> None:
    path = tmp_path / "internal_chain.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Model")
    for row in range(3, 8):
        ws.write_number(row - 1, 0, float(row))
        ws.write_formula(row - 1, 1, f"=A{row}*2")
        ws.write_formula(row - 1, 2, f"=B{row}+1")
        ws.write_formula(row - 1, 3, f"=C{row}")
    wb.close()

    graph = create_dependency_graph(path, ["Model!D3:D7"], load_values=False)
    generator = CodeGenerator(graph)
    config = generator.build_codegen_taco_config(
        ["Model!D3:D7"],
        input_ranges=[f"Model!A{row}" for row in range(3, 8)],
    )
    index = generator.build_codegen_taco_index(
        ["Model!D3:D7"],
        input_ranges=[f"Model!A{row}" for row in range(3, 8)],
    )
    assert_codegen_index_boundaries(graph, index, config)
    assert graph.codegen_taco_index is index
    rr = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr) == 1
    assert rr[0].dependent.min_col == "C"
    assert rr[0].dependent.min_col != "D"


def test_targets_never_appear_on_compressed_dependent_side() -> None:
    graph = DependencyGraph()
    for row in range(3, 8):
        graph.add_node(_make_node(f"Sheet1!B{row}", formula=None, is_leaf=True))
        graph.add_node(
            _make_node(
                f"Sheet1!D{row}",
                formula=f"=B{row}",
                is_target=True,
            ),
        )
        graph.add_edge(f"Sheet1!D{row}", f"Sheet1!B{row}")

    config = TacoBuildConfig.for_codegen(graph)
    index = build_taco_index(graph, config)
    assert_codegen_index_boundaries(graph, index, config)
    assert not index.compressed_edges
