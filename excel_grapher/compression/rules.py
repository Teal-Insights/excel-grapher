"""Compression rule metadata."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass

from .stats import CompressionStats
from .types import CompressedNode

RuleApplyFn = Callable[[Mapping[str, CompressedNode], CompressionStats], dict[str, CompressedNode]]


@dataclass(frozen=True, slots=True)
class RuleSpec:
    """Metadata for one compression rule in pipeline order."""

    rule_id: str
    name: str
    reduces_emission_units: bool
    apply: RuleApplyFn | None = None


COMPRESSION_RULES: tuple[RuleSpec, ...] = (
    RuleSpec("pass_through", "Direct Cell Reference Elimination", False),
    RuleSpec("parallel_if_row", "Parallel Row Compression", True),
    RuleSpec("constant_folding", "Constant Folding", False),
    RuleSpec("common_subexpression", "Common Subexpression Elimination", False),
    RuleSpec("taco_rr", "TACO RR (Relative-Relative)", True),
    RuleSpec("taco_rf", "TACO RF (Relative-Fixed)", True),
    RuleSpec("taco_fr", "TACO FR (Fixed-Relative)", True),
    RuleSpec("taco_ff", "TACO FF (Fixed-Fixed)", True),
    RuleSpec("taco_rr_chain", "TACO RR-Chain", True),
)


def compression_rule_ids() -> list[str]:
    """Return compression rule ids in pipeline order."""
    return [rule.rule_id for rule in COMPRESSION_RULES]
