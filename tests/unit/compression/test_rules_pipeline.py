"""Unit tests for compression rule pipeline wiring."""

from __future__ import annotations

from excel_grapher.compression import (
    apply_compression_rules,
    compression_rules_with_apply,
    expand_compressed_to_cells,
    get_rule_apply,
)
from excel_grapher.compression.nodes import ParallelFormulaNode
from excel_grapher.compression.parity import assert_compression_parity
from excel_grapher.compression.stats import empty_compression_stats
from excel_grapher.core.formula_ast import CellRefNode, NumberNode

from .conftest import parse_formula


def test_get_rule_apply_wires_rules_one_two_and_three() -> None:
    assert get_rule_apply("pass_through") is not None
    assert get_rule_apply("parallel_if_row") is not None
    assert get_rule_apply("constant_folding") is not None


def test_compression_rules_with_apply_populates_apply_fields() -> None:
    by_id = {rule.rule_id: rule for rule in compression_rules_with_apply()}
    assert by_id["pass_through"].apply is not None
    assert by_id["parallel_if_row"].apply is not None
    assert by_id["constant_folding"].apply is not None


def test_apply_compression_rules_runs_pass_through_parallel_then_constant_folding() -> None:
    original = {
        "Sheet1!A1": CellRefNode("Sheet1!B1"),
        "Sheet1!C1": parse_formula("=Sheet1!A1+10"),
        "Sheet1!E1": parse_formula("=2+3+Sheet1!D1"),
    }
    stats = empty_compression_stats()
    compressed = apply_compression_rules(original, stats=stats)

    assert compressed["Sheet1!C1"] == parse_formula("=Sheet1!B1+10")
    assert compressed["Sheet1!E1"] == parse_formula("=5+Sheet1!D1")

    pass_through = stats.contribution_for("pass_through")
    folding = stats.contribution_for("constant_folding")
    assert pass_through.in_place_transforms == 1
    assert folding.in_place_transforms == 1


def test_apply_compression_rules_expand_and_parity() -> None:
    input_values = {"Sheet1!B1": 42, "Sheet1!D1": 2}
    original = {
        "Sheet1!A1": CellRefNode("Sheet1!B1"),
        "Sheet1!C1": parse_formula("=Sheet1!A1+10"),
        "Sheet1!E1": parse_formula("=2+3+Sheet1!D1"),
    }
    compressed = apply_compression_rules(original)
    expanded = expand_compressed_to_cells(compressed)
    assert expanded == {
        key: node for key, node in compressed.items() if not isinstance(node, ParallelFormulaNode)
    }
    assert_compression_parity(original, compressed, input_values=input_values)


def test_apply_compression_rules_respects_explicit_rule_ids() -> None:
    original = {"Sheet1!A1": parse_formula("=2+3")}
    folded_only = apply_compression_rules(original, rule_ids=["constant_folding"])
    assert folded_only["Sheet1!A1"] == NumberNode(5.0)

    unchanged = apply_compression_rules(original, rule_ids=["pass_through"])
    assert unchanged["Sheet1!A1"] == parse_formula("=2+3")
