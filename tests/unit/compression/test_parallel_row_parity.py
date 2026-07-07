"""Pipeline parity tests for parallel row compression."""

from __future__ import annotations

from excel_grapher.compression import (
    apply_compression_rules,
    expand_compressed_to_cells,
)
from excel_grapher.compression.nodes import ParallelFormulaNode
from excel_grapher.compression.parity import assert_compression_parity
from excel_grapher.compression.stats import empty_compression_stats
from excel_grapher.core.address_keys import format_cell_key
from excel_grapher.core.formula_ast import CellRefNode, NumberNode
from excel_grapher.core.types import CellValue

from .conftest import parse_formula
from .test_parallel_row import _if_row_cells, _row_cell


def _pipeline_input_values() -> dict[str, CellValue]:
    return {
        "Sheet1!B1": 42,
        "Sheet1!D1": 2,
        "Ext!D3": "Yes",
        "Ext!D87": 10,
        "Ext!E87": 20,
        "Ext!F87": 30,
        "Sheet1!D10": 1,
        "Sheet1!E10": 2,
        "Sheet1!F10": 3,
    }


def _combined_workbook_map():
    parallel = {cell.key: cell.ast for cell in _if_row_cells()}
    return {
        "Sheet1!A1": CellRefNode("Sheet1!B1"),
        "Sheet1!C1": parse_formula("=Sheet1!A1+10"),
        "Sheet1!E1": parse_formula("=2+3+Sheet1!D1"),
        "Sheet1!Z1": parse_formula("=2+3"),
        **parallel,
    }


def test_pipeline_default_order_runs_one_two_three() -> None:
    stats = empty_compression_stats()
    compressed = apply_compression_rules(_combined_workbook_map(), stats=stats)
    assert compressed["Sheet1!C1"] == parse_formula("=Sheet1!B1+10")
    assert compressed["Sheet1!E1"] == parse_formula("=5+Sheet1!D1")
    assert compressed["Sheet1!Z1"] == NumberNode(5.0)
    assert not any(format_cell_key("Chart Data", col, 177) in compressed for col in ("D", "E", "F"))
    assert any(isinstance(node, ParallelFormulaNode) for node in compressed.values())
    assert stats.contribution_for("pass_through").in_place_transforms == 1
    assert stats.contribution_for("parallel_if_row").cells_affected == 3
    assert stats.contribution_for("constant_folding").in_place_transforms == 2


def test_pipeline_rules_one_three_two_explicit_order() -> None:
    compressed = apply_compression_rules(
        _combined_workbook_map(),
        rule_ids=["pass_through", "constant_folding", "parallel_if_row"],
    )
    assert compressed["Sheet1!Z1"] == NumberNode(5.0)
    assert any(isinstance(node, ParallelFormulaNode) for node in compressed.values())


def test_pipeline_expand_parity_combined_workbook() -> None:
    original = _combined_workbook_map()
    compressed = apply_compression_rules(original)
    expanded = expand_compressed_to_cells(compressed)
    assert set(expanded) == set(original)
    assert_compression_parity(original, compressed, input_values=_pipeline_input_values())


def test_pipeline_preserves_existing_parallel_artifacts() -> None:
    original = {cell.key: cell.ast for cell in _if_row_cells()}
    first_pass = apply_compression_rules(original, rule_ids=["parallel_if_row"])
    artifact_key = next(
        key for key, node in first_pass.items() if isinstance(node, ParallelFormulaNode)
    )
    second_pass = apply_compression_rules(
        first_pass,
        rule_ids=["pass_through", "constant_folding"],
    )
    assert isinstance(second_pass[artifact_key], ParallelFormulaNode)
    assert_compression_parity(original, second_pass, input_values=_pipeline_input_values())


def test_pipeline_does_not_merge_non_contiguous_parallel_columns() -> None:
    ast_map = {
        format_cell_key("Sheet1", col, 10): parse_formula(f"=Sheet1!{col}10*2")
        for col in ("D", "E", "G", "H", "I")
    }
    compressed = apply_compression_rules(ast_map, rule_ids=["parallel_if_row"])
    parallel_nodes = [node for node in compressed.values() if isinstance(node, ParallelFormulaNode)]
    assert len(parallel_nodes) == 1
    node = parallel_nodes[0]
    assert node.start_col == "G"
    assert node.end_col == "I"
    assert format_cell_key("Sheet1", "D", 10) in compressed
    assert format_cell_key("Sheet1", "E", 10) in compressed


def test_pipeline_does_not_merge_mismatched_row_operands() -> None:
    cells = [
        _row_cell("Sheet1", "D", 1, "=Sheet1!D10*2"),
        _row_cell("Sheet1", "E", 1, "=Sheet1!E10*2"),
        _row_cell("Sheet1", "F", 1, "=Sheet1!F10*3"),
    ]
    original = {cell.key: cell.ast for cell in cells}
    compressed = apply_compression_rules(original, rule_ids=["parallel_if_row"])
    assert len(compressed) == 3
    assert not any(isinstance(node, ParallelFormulaNode) for node in compressed.values())
