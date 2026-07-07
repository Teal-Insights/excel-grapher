"""Smoke tests for the compression package scaffold."""

from __future__ import annotations

from excel_grapher.compression import (
    COMPRESSION_RULES,
    ColumnVarCellRefNode,
    CompressionStats,
    ParallelFormulaNode,
    RuleContribution,
    RuleSpec,
    SubexpressionRefNode,
    TacoPatternNode,
    assert_compression_parity,
    expand_compressed_to_cells,
)
from excel_grapher.compression.types import CompressedNode


def test_compression_package_exports_public_api() -> None:
    assert ColumnVarCellRefNode is not None
    assert SubexpressionRefNode is not None
    assert ParallelFormulaNode is not None
    assert TacoPatternNode is not None
    assert RuleContribution is not None
    assert CompressionStats is not None
    assert RuleSpec is not None
    assert COMPRESSION_RULES is not None
    assert expand_compressed_to_cells is not None
    assert assert_compression_parity is not None
    assert CompressedNode is not None


def test_compression_rules_lists_nine_rules_in_pipeline_order() -> None:
    rule_ids = [rule.rule_id for rule in COMPRESSION_RULES]
    assert rule_ids == [
        "pass_through",
        "parallel_if_row",
        "constant_folding",
        "common_subexpression",
        "taco_rr",
        "taco_rf",
        "taco_fr",
        "taco_ff",
        "taco_rr_chain",
    ]


def test_expand_empty_map_returns_no_cells() -> None:
    assert expand_compressed_to_cells({}) == {}
