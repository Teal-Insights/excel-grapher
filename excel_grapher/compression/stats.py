"""Compression statistics and per-rule contribution tracking."""

from __future__ import annotations

from dataclasses import dataclass, field


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

    def record(
        self,
        *,
        cells_affected: int = 0,
        emission_units_saved: int = 0,
        in_place_transforms: int = 0,
        binding_sites: int = 0,
        redundant_evaluations_eliminated: int = 0,
        ast_subnodes_saved: int = 0,
        candidates_rejected: int = 0,
    ) -> None:
        """Increment one or more contribution counters."""
        self.cells_affected += cells_affected
        self.emission_units_saved += emission_units_saved
        self.in_place_transforms += in_place_transforms
        self.binding_sites += binding_sites
        self.redundant_evaluations_eliminated += redundant_evaluations_eliminated
        self.ast_subnodes_saved += ast_subnodes_saved
        self.candidates_rejected += candidates_rejected


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

    def contribution_for(self, rule_id: str) -> RuleContribution:
        """Return the contribution bucket for `rule_id`, creating it if needed."""
        for contribution in self.rule_contributions:
            if contribution.rule_id == rule_id:
                return contribution
        contribution = RuleContribution(rule_id=rule_id)
        self.rule_contributions.append(contribution)
        return contribution


def empty_compression_stats() -> CompressionStats:
    """Return a fresh stats object with zeroed counters."""
    return CompressionStats()
