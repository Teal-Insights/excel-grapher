"""Evaluator ↔ export parity with ``unpack_return=True`` on shared scenarios."""

from __future__ import annotations

import pytest

from tests.integration.exporter.parity_scenarios import parity_scenarios
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


@pytest.mark.parametrize(
    "scenario",
    parity_scenarios(),
    ids=lambda scenario: scenario.name,
)
def test_codegen_matches_evaluator_with_unpack_return(scenario) -> None:
    assert_codegen_matches_evaluator(
        scenario.graph,
        scenario.targets,
        rtol=scenario.rtol,
        atol=scenario.atol,
        blank_ranges=scenario.blank_ranges,
        unpack_return=True,
    )
