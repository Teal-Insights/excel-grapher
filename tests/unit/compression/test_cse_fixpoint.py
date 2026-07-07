"""Unit tests for CSE fixpoint loop and apply_cell_cse."""

from __future__ import annotations

from excel_grapher.compression import empty_compression_stats
from excel_grapher.compression.cse import (
    apply_cell_cse,
    hoist_common_subexpressions_to_fixpoint,
    hoist_one_subexpression,
)
from excel_grapher.compression.expand import expand_compressed_to_cells
from excel_grapher.compression.nodes import ParallelFormulaNode
from excel_grapher.compression.parallel_row import (
    RowCell,
    build_parallel_node,
    find_parallel_runs,
    parallel_artifact_key,
)
from excel_grapher.compression.parity import assert_compression_parity
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    NumberNode,
    SubexpressionRefNode,
)

from .conftest import parse_formula
from .test_cse_subtree import _shared_sum_times_three
from .test_parallel_row import _if_row_cells


def parallel_artifact_key_from_row(row_cells: list[RowCell]) -> str:
    run = find_parallel_runs(row_cells)[0]
    return parallel_artifact_key(run)


def _shared_sum_times_two_all_three() -> dict[str, AstNode]:
    shared = parse_formula("=Sheet1!B1+Sheet1!C1")
    return {
        "Sheet1!A1": BinaryOpNode("*", shared, NumberNode(2.0)),
        "Sheet1!A2": BinaryOpNode("*", shared, NumberNode(2.0)),
        "Sheet1!A3": BinaryOpNode("*", shared, NumberNode(2.0)),
    }


def _after_first_cse_round_with_ref_times_two() -> dict[str, AstNode]:
    shared = parse_formula("=Sheet1!B1+Sheet1!C1")
    ref_times_two = BinaryOpNode("*", SubexpressionRefNode("_cse!0"), NumberNode(2.0))
    return {
        "_cse!0": shared,
        "Sheet1!A1": ref_times_two,
        "Sheet1!A2": ref_times_two,
        "Sheet1!A3": ref_times_two,
    }


def test_hoist_fixpoint_single_round_matches_one_hoist() -> None:
    cell_map = _shared_sum_times_three()
    one_round, one_result = hoist_one_subexpression(cell_map)
    fixpoint, fixpoint_result = hoist_common_subexpressions_to_fixpoint(cell_map)
    assert fixpoint == one_round
    assert fixpoint_result.cse_fixpoint_rounds == 1
    assert fixpoint_result.binding_sites == one_result.binding_sites
    assert fixpoint_result.hoisted


def test_hoist_fixpoint_hoists_identical_formula_in_one_round() -> None:
    cell_map = _shared_sum_times_two_all_three()
    fixpoint, result = hoist_common_subexpressions_to_fixpoint(cell_map)
    assert result.cse_fixpoint_rounds == 1
    assert result.binding_sites == 1
    assert fixpoint["Sheet1!A1"] == SubexpressionRefNode("_cse!0")
    assert fixpoint["_cse!0"] == BinaryOpNode(
        "*", parse_formula("=Sheet1!B1+Sheet1!C1"), NumberNode(2.0)
    )


def test_hoist_fixpoint_second_round_hoists_over_cse_refs() -> None:
    fixpoint, result = hoist_common_subexpressions_to_fixpoint(
        _after_first_cse_round_with_ref_times_two()
    )
    assert result.cse_fixpoint_rounds == 1
    assert result.binding_sites == 1
    assert "_cse!1" in fixpoint
    assert fixpoint["Sheet1!A1"] == SubexpressionRefNode("_cse!1")
    assert fixpoint["Sheet1!A2"] == SubexpressionRefNode("_cse!1")
    assert fixpoint["Sheet1!A3"] == SubexpressionRefNode("_cse!1")
    assert fixpoint["_cse!1"] == BinaryOpNode("*", SubexpressionRefNode("_cse!0"), NumberNode(2.0))


def test_hoist_fixpoint_expand_parity_after_two_rounds() -> None:
    original = _shared_sum_times_two_all_three()
    compressed, result = hoist_common_subexpressions_to_fixpoint(original)
    assert result.cse_fixpoint_rounds == 1
    assert expand_compressed_to_cells(compressed) == original
    assert_compression_parity(
        original,
        compressed,
        input_values={"Sheet1!B1": 2, "Sheet1!C1": 3},
    )

    second_pass, second_result = hoist_common_subexpressions_to_fixpoint(compressed)
    assert second_result.cse_fixpoint_rounds == 0
    assert second_pass == compressed


def test_apply_cell_cse_on_plain_cell_map() -> None:
    original = _shared_sum_times_three()
    result = apply_cell_cse(original)
    assert "_cse!0" in result
    assert expand_compressed_to_cells(result) == original


def test_apply_cell_cse_on_mixed_map_preserves_parallel_artifact() -> None:
    row_cells = _if_row_cells()
    cse_cells = _shared_sum_times_three()
    compressed = {
        **cse_cells,
        parallel_artifact_key_from_row(row_cells): build_parallel_node(
            find_parallel_runs(row_cells)[0]
        ),
    }
    original_cells = {
        **cse_cells,
        **{cell.key: cell.ast for cell in row_cells},
    }
    result = apply_cell_cse(compressed)
    assert any(isinstance(value, ParallelFormulaNode) for value in result.values())
    assert "_cse!0" in result
    expanded = expand_compressed_to_cells(result)
    assert set(expanded) == set(original_cells)
    assert_compression_parity(
        original_cells,
        result,
        input_values={
            "Sheet1!B1": 2,
            "Sheet1!C1": 3,
            "Ext!D3": "Yes",
            "Ext!D87": 10,
            "Ext!E87": 20,
            "Ext!F87": 30,
        },
    )


def test_apply_cell_cse_records_stats() -> None:
    stats = empty_compression_stats()
    apply_cell_cse(_shared_sum_times_three(), stats=stats)
    assert stats.cse_fixpoint_rounds == 1
    assert stats.binding_sites == 1
    assert stats.ast_subnodes_saved == 4
    assert stats.redundant_evaluations_eliminated == 2
    contribution = stats.contribution_for("common_subexpression")
    assert contribution.binding_sites == 1
    assert contribution.ast_subnodes_saved == 4
    assert contribution.candidates_rejected >= 0
