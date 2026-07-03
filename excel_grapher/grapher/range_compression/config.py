"""Configuration for TACO index construction."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey


@dataclass(frozen=True, slots=True)
class TacoBuildConfig:
    """Controls which cells may participate in TACO range-pattern compression.

    For codegen, start with ``exclude_targets`` and ``exclude_input_keys`` so
    boundary cells stay at cell granularity. Use ``internal_only`` to compress
    only formula nodes that are neither targets nor declared inputs.
    """

    exclude_targets: bool = False
    exclude_input_keys: frozenset[NodeKey] = frozenset()
    internal_only: bool = False

    @classmethod
    def for_codegen(
        cls,
        graph: DependencyGraph | None = None,
        *,
        input_keys: frozenset[NodeKey] | None = None,
        internal_only: bool = True,
    ) -> TacoBuildConfig:
        """Preset for codegen: keep targets and inputs uncompressed."""
        if input_keys is None and graph is not None:
            input_keys = input_keys_from_graph(graph)
        return cls(
            exclude_targets=True,
            exclude_input_keys=input_keys or frozenset(),
            internal_only=internal_only,
        )


def input_keys_from_graph(graph: DependencyGraph) -> frozenset[NodeKey]:
    """Return sheet-qualified keys marked ``input`` in ``graph.leaf_classification``."""
    lc = graph.leaf_classification
    if not lc:
        return frozenset()
    return frozenset(k for k, role in lc.items() if role == "input")
