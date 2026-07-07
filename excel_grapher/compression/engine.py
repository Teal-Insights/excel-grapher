"""Compression engine orchestration for rule pipelines."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from functools import lru_cache

from excel_grapher.core.formula_ast import AstNode

from .rules import COMPRESSION_RULES, RuleApplyFn, RuleSpec
from .stats import CompressionStats, empty_compression_stats
from .types import CompressedNode, normalize_compressed_key


@lru_cache(maxsize=1)
def _implemented_rule_appliers() -> dict[str, RuleApplyFn]:
    from .constant_folding import apply_constant_folding
    from .parallel_row import apply_parallel_row
    from .pass_through import apply_pass_through

    return {
        "pass_through": apply_pass_through,
        "parallel_if_row": apply_parallel_row,
        "constant_folding": apply_constant_folding,
    }


def get_rule_apply(rule_id: str) -> RuleApplyFn | None:
    """Return the apply function for an implemented rule id."""
    return _implemented_rule_appliers().get(rule_id)


def compression_rules_with_apply() -> tuple[RuleSpec, ...]:
    """Return rule metadata with `apply` wired for implemented rules."""
    appliers = _implemented_rule_appliers()
    return tuple(
        RuleSpec(rule.rule_id, rule.name, rule.reduces_emission_units, appliers.get(rule.rule_id))
        for rule in COMPRESSION_RULES
    )


def apply_compression_rules(
    ast_map: Mapping[str, AstNode] | Mapping[str, CompressedNode],
    *,
    rule_ids: Sequence[str] | None = None,
    stats: CompressionStats | None = None,
) -> dict[str, CompressedNode]:
    """Apply compression rules to a per-cell AST map in pipeline order.

    Args:
        ast_map: Sheet-qualified cell keys mapped to formula ASTs or compressed
            nodes from an earlier pipeline stage.
        rule_ids: Rule ids to run, in the order given. Defaults to all
            implemented rules in pipeline order.
        stats: Optional stats object to accumulate rule contributions.

    Returns:
        Mixed compressed map after the selected rules run.
    """
    stats_obj = stats if stats is not None else empty_compression_stats()
    working: dict[str, CompressedNode] = {
        normalize_compressed_key(cell_key): node for cell_key, node in ast_map.items()
    }
    appliers = _implemented_rule_appliers()
    selected_rule_ids = (
        list(rule_ids)
        if rule_ids is not None
        else [rule.rule_id for rule in COMPRESSION_RULES if rule.rule_id in appliers]
    )

    for rule_id in selected_rule_ids:
        apply_fn = appliers.get(rule_id)
        if apply_fn is None:
            continue
        working = apply_fn(working, stats_obj)
    return working
