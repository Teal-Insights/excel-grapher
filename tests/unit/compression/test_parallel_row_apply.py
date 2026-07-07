"""Unit tests for applying parallel row compression."""

from __future__ import annotations

from excel_grapher.compression import empty_compression_stats
from excel_grapher.compression.expand import expand_compressed_to_cells, materialize_parallel_node
from excel_grapher.compression.nodes import ParallelFormulaNode
from excel_grapher.compression.parallel_row import (
    apply_parallel_row,
    build_parallel_node,
    find_parallel_runs,
    parallel_artifact_key,
)
from excel_grapher.compression.parity import assert_compression_parity
from excel_grapher.core.address_keys import format_cell_key
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    ColumnVarCellRefNode,
    FunctionCallNode,
    NumberNode,
)

from .conftest import parse_formula
from .test_parallel_row import _if_row_cells, _row_cell


def test_build_parallel_node_if_row() -> None:
    run = find_parallel_runs(_if_row_cells())[0]
    node = build_parallel_node(run)
    assert node.sheet == "Chart Data"
    assert node.start_col == "D"
    assert node.end_col == "F"
    assert node.output_row == 177
    assert isinstance(node.template, FunctionCallNode)
    assert node.template.name.upper() == "IF"
    assert node.condition == node.template.args[0]
    assert node.if_true == node.template.args[1]
    assert node.if_false == node.template.args[2]
    assert node.if_false == ColumnVarCellRefNode(column_variable="COL", sheet="Ext", row=87)


def test_parallel_artifact_key_includes_bounds() -> None:
    run = find_parallel_runs(_if_row_cells())[0]
    assert parallel_artifact_key(run) == "parallel:'Chart Data'!177:D:F"


def test_apply_parallel_row_removes_absorbed_cells() -> None:
    original = {cell.key: cell.ast for cell in _if_row_cells()}
    original[format_cell_key("Sheet1", "A", 1)] = parse_formula("=1+1")
    compressed = apply_parallel_row(original)
    assert format_cell_key("Chart Data", "D", 177) not in compressed
    assert format_cell_key("Chart Data", "E", 177) not in compressed
    assert format_cell_key("Chart Data", "F", 177) not in compressed
    assert format_cell_key("Sheet1", "A", 1) in compressed
    artifact_keys = [
        key for key, value in compressed.items() if isinstance(value, ParallelFormulaNode)
    ]
    assert len(artifact_keys) == 1


def test_materialize_parallel_node_matches_originals() -> None:
    original = {cell.key: cell.ast for cell in _if_row_cells()}
    compressed = apply_parallel_row(original)
    artifact = next(node for node in compressed.values() if isinstance(node, ParallelFormulaNode))
    materialized = materialize_parallel_node(artifact)
    assert set(materialized) == set(original)
    assert materialized == original


def test_apply_parallel_row_expand_parity_if_row() -> None:
    original = {cell.key: cell.ast for cell in _if_row_cells()}
    compressed = apply_parallel_row(original)
    expanded = expand_compressed_to_cells(compressed)
    assert expanded == original
    assert_compression_parity(
        original,
        compressed,
        input_values={
            "Ext!D3": "Yes",
            "Ext!D87": 10,
            "Ext!E87": 20,
            "Ext!F87": 30,
        },
    )


def test_apply_parallel_row_expand_parity_multiply_row() -> None:
    cells = [_row_cell("Sheet1", col, 1, f"=Sheet1!{col}10*2") for col in ("D", "E", "F")]
    original = {cell.key: cell.ast for cell in cells}
    compressed = apply_parallel_row(original)
    expanded = expand_compressed_to_cells(compressed)
    assert expanded == original
    assert_compression_parity(
        original,
        compressed,
        input_values={
            "Sheet1!D10": 1,
            "Sheet1!E10": 2,
            "Sheet1!F10": 3,
        },
    )


def test_apply_parallel_row_leaves_length_two_runs() -> None:
    cells = [
        _row_cell("Sheet1", "D", 1, "=Sheet1!D10*2"),
        _row_cell("Sheet1", "E", 1, "=Sheet1!E10*2"),
    ]
    original = {cell.key: cell.ast for cell in cells}
    compressed = apply_parallel_row(original)
    assert len(compressed) == 2
    assert all(not isinstance(node, ParallelFormulaNode) for node in compressed.values())


def test_apply_parallel_row_records_stats() -> None:
    original = {cell.key: cell.ast for cell in _if_row_cells()}
    stats = empty_compression_stats()
    apply_parallel_row(original, stats)
    contribution = stats.contribution_for("parallel_if_row")
    assert contribution.cells_affected == 3
    assert contribution.emission_units_saved == 2


def test_build_parallel_node_non_if_has_no_projection() -> None:
    cells = [_row_cell("Sheet1", col, 1, f"=Sheet1!{col}10*2") for col in ("D", "E", "F")]
    node = build_parallel_node(find_parallel_runs(cells)[0])
    assert node.condition is None
    assert node.if_true is None
    assert node.if_false is None
    assert node.template == BinaryOpNode(
        "*",
        ColumnVarCellRefNode(column_variable="COL", sheet="Sheet1", row=10),
        NumberNode(2.0),
    )


def test_apply_parallel_row_mixed_fixed_and_parallel_cells() -> None:
    parallel = {cell.key: cell.ast for cell in _if_row_cells()}
    solo = {format_cell_key("Sheet1", "A", 1): parse_formula("=2+3")}
    original = {**parallel, **solo}
    compressed = apply_parallel_row(original)
    assert format_cell_key("Sheet1", "A", 1) in compressed
    assert compressed[format_cell_key("Sheet1", "A", 1)] == parse_formula("=2+3")
