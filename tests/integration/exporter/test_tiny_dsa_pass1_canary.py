"""Tiny DSA Pass-1 canary (issue #595) — shape RED; numeric parity stays green.

Vendors the Tiny DSA workbook + bindings and asserts the desired Pass-1 helper
surface (inventory from bindings, no ``cell_*`` for bound formula series,
self-recurrence, auto-wired ``compute_*``). Shape contracts fail until
series-helper collapse lands. A separate numeric parity test must keep passing
on today's ``cell_*`` export.
"""

from __future__ import annotations

import json
from typing import Any, cast

import pytest

from excel_grapher import DynamicRefConfig, FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, load_series_bindings
from tests.fixtures.tiny_dsa.constraints import BLANK_RANGES, CONSTRAINTS
from tests.integration.exporter.pass1_shape_contract import (
    assert_compute_calls_helper,
    assert_helper_inventory,
    assert_helper_signature,
    assert_no_cell_defs_for_addresses,
    bound_formula_addresses_from_bindings,
    expected_helper_signature,
    formula_series_ids_from_bindings,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.paths import TINY_DSA_FIXTURES

WORKBOOK = TINY_DSA_FIXTURES / "tiny-dsa.xlsx"
BINDINGS_DIR = TINY_DSA_FIXTURES
TARGETS_PATH = TINY_DSA_FIXTURES / "targets.json"

# Flagship helpers called out in #595 Layer B.
_TIME_KEYED_HELPERS = frozenset(
    {
        "shock_active",
        "shocked_growth",
        "shocked_interest",
        "shocked_primary_balance",
        "baseline_path_internal",
        "shocked_path_internal",
        "output_baseline",
        "output_shocked",
        "output_delta",
    }
)
_SCALAR_HELPERS = frozenset(
    {
        "initial_debt_resolved",
        "engine_initial_debt_baseline",
        "engine_initial_debt_shocked",
        "shock_magnitude_resolved",
    }
)


@pytest.fixture(scope="module")
def tiny_dsa_bindings():
    return load_series_bindings(BINDINGS_DIR)


@pytest.fixture(scope="module")
def tiny_dsa_targets() -> list[str]:
    raw = json.loads(TARGETS_PATH.read_text(encoding="utf-8"))
    assert isinstance(raw, list)
    return [str(t) for t in raw]


@pytest.fixture(scope="module")
def tiny_dsa_graph(tiny_dsa_targets: list[str]):
    return create_dependency_graph(
        WORKBOOK,
        tiny_dsa_targets,
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_constraints(CONSTRAINTS, {}),
        blank_ranges=BLANK_RANGES,
    )


@pytest.fixture(scope="module")
def tiny_dsa_modules(tiny_dsa_graph, tiny_dsa_bindings, tiny_dsa_targets: list[str]):
    # Expand named-range targets to concrete cells for generate_modules.
    expanded: list[str] = []
    for target in tiny_dsa_targets:
        expanded.extend(expand_data_range(target, workbook=WORKBOOK))
    with CodeGenerator(tiny_dsa_graph) as gen:
        files = gen.generate_modules(
            expanded,
            series_bindings=tiny_dsa_bindings,
            bindings_workbook=WORKBOOK,
            blank_ranges=BLANK_RANGES,
        )
    return files, expanded


def _formula_addresses(graph) -> set[str]:
    out: set[str] = set()
    for key in graph:
        node = graph.get_node(key)
        if node is not None and node.has_formula:
            out.add(str(key))
    return out


def test_tiny_dsa_pass1_helper_shape_contract(
    tiny_dsa_modules,
    tiny_dsa_bindings,
    tiny_dsa_graph,
) -> None:
    """Pass-1 shape canary — RED until bound series collapse (#595)."""
    files, _expanded = tiny_dsa_modules
    internals = files["internals.py"]
    api = files["api.py"]
    formula_addrs = _formula_addresses(tiny_dsa_graph)

    expected_ids = formula_series_ids_from_bindings(
        cast(dict[str, Any], tiny_dsa_bindings),
        workbook=WORKBOOK,
        formula_addresses=formula_addrs,
    )
    # v1: every formula-backed internal/output binding gets a helper — including
    # series today's pipeline left as cell_* wrappers.
    assert expected_ids >= (_TIME_KEYED_HELPERS | _SCALAR_HELPERS)

    assert_helper_inventory(internals, expected_ids)

    series_by_id = {s["id"]: s for s in tiny_dsa_bindings["series"] if isinstance(s.get("id"), str)}
    for series_id in expected_ids:
        series = series_by_id[series_id]
        assert_helper_signature(internals, series_id, expected_helper_signature(series))

    bound_addrs = bound_formula_addresses_from_bindings(
        cast(dict[str, Any], tiny_dsa_bindings),
        workbook=WORKBOOK,
        formula_addresses=formula_addrs,
    )
    assert_no_cell_defs_for_addresses(internals, bound_addrs)

    for name in ("baseline_path_internal", "shocked_path_internal"):
        body = internals  # search whole module; helpers must contain recurrence
        assert "time_period - 1" in body or "time_period-1" in body, name
        assert f"{name}(ctx, time_period=" in body or f"def {name}(" in body

    # shock_active is one parameterized helper, not five cell_* year dumps.
    assert "def shock_active(" in internals
    assert "def cell_engine_c10(" not in internals

    for series_id, compute_name in (
        ("output_baseline", "compute_output_baseline"),
        ("output_shocked", "compute_output_shocked"),
        ("output_delta", "compute_output_delta"),
    ):
        output_addrs = expand_data_range(
            series_by_id[series_id]["data_range"],
            workbook=WORKBOOK,
        )
        assert_compute_calls_helper(
            api,
            compute_name,
            series_id,
            output_addresses=output_addrs,
        )


def test_tiny_dsa_numeric_parity_still_passes(
    tiny_dsa_graph,
    tiny_dsa_targets: list[str],
) -> None:
    """Numeric fidelity of today's cell_* export must remain green (#595)."""
    expanded: list[str] = []
    for target in tiny_dsa_targets:
        expanded.extend(expand_data_range(target, workbook=WORKBOOK))

    # Smoke: evaluator can resolve the output strip under constraints.
    evaluator = FormulaEvaluator(
        tiny_dsa_graph,
        blank_ranges=BLANK_RANGES,
    )
    for addr in expanded:
        evaluator.evaluate(addr)

    assert_codegen_matches_evaluator(
        tiny_dsa_graph,
        expanded,
        blank_ranges=BLANK_RANGES,
    )
