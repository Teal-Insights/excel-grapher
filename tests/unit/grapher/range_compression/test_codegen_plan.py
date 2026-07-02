"""Codegen evaluation plan over codegen-boundary TACO indexes."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher.range_compression import (
    CompressedUnit,
    PatternKind,
    TacoBuildConfig,
    build_codegen_plan,
    build_taco_index,
    range_ref_unit_id,
)


def _internal_chain_workbook(tmp_path: Path) -> Path:
    path = tmp_path / "internal_chain.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Model")
    for row in range(3, 8):
        ws.write_number(row - 1, 0, float(row))
        ws.write_formula(row - 1, 1, f"=A{row}*2")
        ws.write_formula(row - 1, 2, f"=B{row}+1")
        ws.write_formula(row - 1, 3, f"=C{row}")
    wb.close()
    return path


def _formula_closure(graph, targets: list[str]) -> list[str]:
    roots = list(graph.target_keys()) if not targets else list(targets)
    stack = list(roots)
    seen = set(roots)
    while stack:
        key = stack.pop()
        for dep in graph.get_dependencies(key):
            if dep not in seen:
                seen.add(dep)
                stack.append(dep)
    return sorted(key for key in seen if (node := graph.get_node(key)) is not None and node.formula)


def test_range_ref_unit_id_column_span() -> None:
    from excel_grapher.grapher.range_compression import RangeRef

    ref = RangeRef.column_span("Model", "C", 3, 7)
    assert range_ref_unit_id(ref) == "Model!C3:C7"


def test_internal_chain_plan_has_compressed_and_single_units(tmp_path: Path) -> None:
    path = _internal_chain_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Model!D3:D7"], load_values=False)
    inputs = frozenset(f"Model!A{row}" for row in range(3, 8))
    config = TacoBuildConfig.for_codegen(input_keys=inputs)
    index = build_taco_index(graph, config)
    closure = _formula_closure(graph, [])

    plan = build_codegen_plan(graph, index, config, closure=closure)

    assert len(plan.compressed_units) == 1
    unit = plan.compressed_units[0]
    assert isinstance(unit, CompressedUnit)
    assert unit.unit_id == "Model!C3:C7"
    assert unit.edge.meta.kind == PatternKind.rr
    assert set(plan.single_cells) == {f"Model!B{row}" for row in range(3, 8)} | {
        f"Model!D{row}" for row in range(3, 8)
    }


def test_internal_chain_cell_to_unit_coverage(tmp_path: Path) -> None:
    path = _internal_chain_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Model!D3:D7"], load_values=False)
    inputs = frozenset(f"Model!A{row}" for row in range(3, 8))
    config = TacoBuildConfig.for_codegen(input_keys=inputs)
    index = build_taco_index(graph, config)
    closure = _formula_closure(graph, [])

    plan = build_codegen_plan(graph, index, config, closure=closure)

    for row in range(3, 8):
        assert plan.cell_to_unit[f"Model!C{row}"] == "Model!C3:C7"
        assert plan.cell_to_unit[f"Model!B{row}"] == f"Model!B{row}"
        assert plan.cell_to_unit[f"Model!D{row}"] == f"Model!D{row}"
        assert plan.cell_to_unit[f"Model!A{row}"] == f"Model!A{row}"


def test_internal_chain_eval_order_respects_dependencies(tmp_path: Path) -> None:
    path = _internal_chain_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Model!D3:D7"], load_values=False)
    inputs = frozenset(f"Model!A{row}" for row in range(3, 8))
    config = TacoBuildConfig.for_codegen(input_keys=inputs)
    index = build_taco_index(graph, config)
    closure = _formula_closure(graph, [])

    plan = build_codegen_plan(graph, index, config, closure=closure)
    order_ids = [unit.unit_id for unit in plan.eval_order]

    c_pos = order_ids.index("Model!C3:C7")
    for row in range(3, 8):
        assert order_ids.index(f"Model!B{row}") < c_pos
        assert order_ids.index(f"Model!D{row}") > c_pos


def test_targets_not_in_compressed_units(tmp_path: Path) -> None:
    path = _internal_chain_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Model!D3:D7"], load_values=False)
    inputs = frozenset(f"Model!A{row}" for row in range(3, 8))
    config = TacoBuildConfig.for_codegen(input_keys=inputs)
    index = build_taco_index(graph, config)
    closure = _formula_closure(graph, [])

    plan = build_codegen_plan(graph, index, config, closure=closure)

    compressed_cells = {key for unit in plan.compressed_units for key in unit.dependent.cell_keys()}
    for row in range(3, 8):
        assert f"Model!D{row}" not in compressed_cells


def test_code_generator_build_codegen_plan(tmp_path: Path) -> None:
    path = _internal_chain_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Model!D3:D7"], load_values=False)
    generator = CodeGenerator(graph)
    plan = generator.build_codegen_plan(
        ["Model!D3:D7"],
        input_ranges=[f"Model!A{row}" for row in range(3, 8)],
    )
    assert len(plan.compressed_units) == 1
    assert plan.compressed_units[0].unit_id == "Model!C3:C7"
    assert graph.codegen_taco_index is plan.index


def test_plan_rejects_cycle_in_units() -> None:
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.node import Node

    graph = DependencyGraph()

    def node(key: str, formula: str) -> Node:
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
            is_leaf=False,
        )

    graph.add_node(node("Sheet1!A1", "=B1"))
    graph.add_node(node("Sheet1!B1", "=A1"))
    graph.add_edge("Sheet1!A1", "Sheet1!B1")
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    index = build_taco_index(graph)
    config = TacoBuildConfig.for_codegen()
    with pytest.raises(ValueError, match="cycle"):
        build_codegen_plan(graph, index, config, closure=["Sheet1!A1", "Sheet1!B1"])
