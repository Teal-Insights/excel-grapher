"""Unit tests for compression rule metadata and statistics stubs."""

from __future__ import annotations

from excel_grapher.compression.engine import get_rule_apply
from excel_grapher.compression.rules import (
    COMPRESSION_RULES,
    compression_rule_ids,
)
from excel_grapher.compression.stats import RuleContribution, empty_compression_stats


def test_compression_rule_ids_match_pipeline_order() -> None:
    assert compression_rule_ids() == [
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
    assert [rule.rule_id for rule in COMPRESSION_RULES] == compression_rule_ids()


def test_compression_rules_metadata_fields() -> None:
    parallel = COMPRESSION_RULES[1]
    assert parallel.rule_id == "parallel_if_row"
    assert parallel.name == "Parallel Row Compression"
    assert parallel.reduces_emission_units is True
    assert parallel.apply is None

    pass_through = COMPRESSION_RULES[0]
    assert pass_through.reduces_emission_units is False
    assert get_rule_apply("pass_through") is not None

    assert get_rule_apply("constant_folding") is not None
    assert get_rule_apply("parallel_if_row") is not None


def test_empty_compression_stats_defaults() -> None:
    stats = empty_compression_stats()
    assert stats.emission_units == 0
    assert stats.binding_sites == 0
    assert stats.rule_contributions == []


def test_rule_contribution_record_increments_counters() -> None:
    contrib = RuleContribution(rule_id="constant_folding")
    contrib.record(in_place_transforms=2, cells_affected=5)
    contrib.record(in_place_transforms=1, ast_subnodes_saved=4)
    assert contrib.in_place_transforms == 3
    assert contrib.cells_affected == 5
    assert contrib.ast_subnodes_saved == 4
    assert contrib.emission_units_saved == 0


def test_compression_stats_contribution_for_creates_entry() -> None:
    stats = empty_compression_stats()
    folding = stats.contribution_for("constant_folding")
    assert folding.rule_id == "constant_folding"
    folding.record(in_place_transforms=1)
    assert stats.contribution_for("constant_folding").in_place_transforms == 1
    assert len(stats.rule_contributions) == 1


def test_compression_stats_contribution_for_reuses_entry() -> None:
    stats = empty_compression_stats()
    first = stats.contribution_for("taco_rr")
    second = stats.contribution_for("taco_rr")
    assert first is second
