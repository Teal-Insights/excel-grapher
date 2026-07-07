"""Compression rule metadata and statistics stubs."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass, field

from .types import CompressedNode


@dataclass(frozen=True, slots=True)
class RuleSpec:
    """Metadata for one compression rule in pipeline order."""

    rule_id: str
    name: str
    reduces_emission_units: bool
    apply: Callable[[Mapping[str, CompressedNode]], None] | None = None


@dataclass
class RuleContribution:
    """Per-rule counters recorded during compression."""

    rule_id: str
    cells_affected: int = 0
    emission_units_saved: int = 0
    in_place_transforms: int = 0
    binding_sites: int = 0
    redundant_evaluations_eliminated: int = 0
    ast_subnodes_saved: int = 0
    candidates_rejected: int = 0


@dataclass
class CompressionStats:
    """Aggregate compression metrics across the full pipeline."""

    emission_units: int = 0
    binding_sites: int = 0
    redundant_evaluations_eliminated: int = 0
    ast_subnodes_saved: int = 0
    dependency_edges: int = 0
    cells_per_emission_unit: float = 0.0
    cse_fixpoint_rounds: int = 0
    artifact_cse_hoists: int = 0
    post_cse_formula_folds: int = 0
    rule_contributions: list[RuleContribution] = field(default_factory=list)


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
