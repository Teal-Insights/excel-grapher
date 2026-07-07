"""Pipeline integration tests for cell-level CSE (rule 4a)."""

from __future__ import annotations

from excel_grapher.compression import (
    apply_compression_rules,
    empty_compression_stats,
    expand_compressed_to_cells,
)
from excel_grapher.compression.nodes import ParallelFormulaNode
from excel_grapher.compression.parallel_row import (
    build_parallel_node,
    find_parallel_runs,
    parallel_artifact_key,
)
from excel_grapher.compression.parity import assert_compression_parity
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    NumberNode,
    SubexpressionRefNode,
)

from .conftest import parse_formula
from .test_cse_subtree import _shared_sum_times_three
from .test_parallel_row import _if_row_cells


def test_pipeline_default_includes_common_subexpression() -> None:
    original = _shared_sum_times_three()
    compressed = apply_compression_rules(original)
    assert "_cse!0" in compressed
    assert expand_compressed_to_cells(compressed) == original


def test_pipeline_rules_one_through_four_cse_parity() -> None:
    original = _shared_sum_times_three()
    compressed = apply_compression_rules(
        original,
        rule_ids=[
            "pass_through",
            "parallel_if_row",
            "constant_folding",
            "common_subexpression",
        ],
    )
    assert "_cse!0" in compressed
    assert isinstance(compressed["Sheet1!A1"], BinaryOpNode)
    assert isinstance(compressed["Sheet1!A1"].left, SubexpressionRefNode)
    assert_compression_parity(
        original,
        compressed,
        input_values={"Sheet1!B1": 2, "Sheet1!C1": 3},
    )


def test_pipeline_cse_with_parallel_artifact_parity() -> None:
    row_cells = _if_row_cells()
    cse_cells = _shared_sum_times_three()
    run = find_parallel_runs(row_cells)[0]
    original = {
        **cse_cells,
        **{cell.key: cell.ast for cell in row_cells},
    }
    compressed = apply_compression_rules(
        {
            **cse_cells,
            parallel_artifact_key(run): build_parallel_node(run),
        },
        rule_ids=["common_subexpression"],
    )
    assert "_cse!0" in compressed
    assert any(isinstance(node, ParallelFormulaNode) for node in compressed.values())
    expanded = expand_compressed_to_cells(compressed)
    assert set(expanded) == set(original)
    assert_compression_parity(
        original,
        compressed,
        input_values={
            "Sheet1!B1": 2,
            "Sheet1!C1": 3,
            "Ext!D3": "Yes",
            "Ext!D87": 10,
            "Ext!E87": 20,
            "Ext!F87": 30,
        },
    )


def test_pipeline_preserves_existing_cse_bindings() -> None:
    shared = parse_formula("=Sheet1!B1+Sheet1!C1")
    existing_binding = parse_formula("=Sheet1!X1+Sheet1!Y1")
    original = {
        "_cse!0": existing_binding,
        "Sheet1!A1": BinaryOpNode("*", shared, NumberNode(2.0)),
        "Sheet1!A2": BinaryOpNode("*", shared, NumberNode(3.0)),
        "Sheet1!A3": BinaryOpNode("+", shared, NumberNode(10.0)),
    }
    compressed = apply_compression_rules(
        original,
        rule_ids=["common_subexpression"],
    )
    assert compressed["_cse!0"] == existing_binding
    assert "_cse!1" in compressed
    assert compressed["_cse!1"] == shared
    assert expand_compressed_to_cells(compressed) == {
        key: ast for key, ast in original.items() if not key.startswith("_cse!")
    }


def test_pipeline_cse_records_stats() -> None:
    stats = empty_compression_stats()
    apply_compression_rules(_shared_sum_times_three(), stats=stats)
    assert stats.cse_fixpoint_rounds == 1
    assert stats.binding_sites == 1
    contribution = stats.contribution_for("common_subexpression")
    assert contribution.binding_sites == 1
    assert contribution.ast_subnodes_saved == 4
