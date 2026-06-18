"""Parity helpers for ``advanced_formula_workbook.xlsx`` binding accuracy tests."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import pytest

from excel_grapher.grapher import DependencyGraph
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from tests.integration.user_flows.bindings_accuracy import resolve_leaves
from tests.integration.user_flows.financial_model_parity import (
    DownstreamChainCase,
    DownstreamChainStep,
    OutputLeaf,
    collect_output_leaves,
    parity_values_equal,
)

SANDBOX = Path(__file__).resolve().parents[3] / "sandbox" / "model"
WORKBOOK = SANDBOX / "advanced_formula_workbook.xlsx"
BINDINGS_DIR = SANDBOX / "advanced_formula_workbook.bindings"

RTOL = 1e-9
ATOL = 1e-9

SPOT_PARITY_SUBGRAPHS: dict[str, tuple[str, ...]] = {
    "xludf_lookups": (
        "'Product Lookup'!K7",
        "'Product Lookup'!K9",
        "'Product Lookup'!K12",
        "'Formula Toolkit'!D12",
        "'Formula Toolkit'!D30",
    ),
    "abs_normdist": tuple(f"'Statistical Analysis'!M{row}" for row in range(19, 23)),
    "lookup_panel": tuple(f"'Product Lookup'!K{row}" for row in range(6, 18)),
    "employee_tenure": tuple(f"'Employee Directory'!I{row}" for row in range(5, 9)),
}

SPOT_OUTPUT_SERIES: frozenset[str] = frozenset(
    {
        "lookup_panel",
        "sumproduct_analytics",
        "normdist_sigma_band",
        "employee_tenure",
    }
)

INPUT_SPOT_ADDRESSES: tuple[str, ...] = (
    "Assumptions!B14",
    "Assumptions!B27",
    "'Product Lookup'!K5",
    "'Product Lookup'!E5",
    "'Statistical Analysis'!D5",
    "'Sales Tracker'!B5",
)

# Evaluator ↔ standalone codegen gaps on spot subgraphs.
XFAIL_EVAL_CODEGEN: frozenset[str] = frozenset(
    {
        "'Product Lookup'!K16",
        "'Product Lookup'!K24",
    }
)

# Modular binding ``compute_*`` vs evaluator gaps (includes partial-graph overlap leaves).
XFAIL_BINDING_EVAL: frozenset[str] = frozenset(
    {
        "'Product Lookup'!K16",
        "'Product Lookup'!K18",
        "'Product Lookup'!K19",
        "'Product Lookup'!K24",
    }
)

# Formula outputs without cached workbook values — eval↔Excel checks are not meaningful.
XFAIL_EVAL_EXCEL: frozenset[str] = frozenset()

XFAIL_CODEGEN_EXCEL: frozenset[str] = frozenset()

ParityLeg = Literal["codegen_excel", "eval_excel", "eval_codegen"]


def maybe_xfail_address(address: str, leg: ParityLeg) -> None:
    """Mark the current test xfail when ``address`` is a known gap for ``leg``."""
    xfail_by_leg = {
        "codegen_excel": XFAIL_CODEGEN_EXCEL,
        "eval_excel": XFAIL_EVAL_EXCEL,
        "eval_codegen": XFAIL_EVAL_CODEGEN,
    }
    if address in xfail_by_leg[leg]:
        pytest.xfail(f"known {leg.replace('_', ' ')} gap at {address}")


def sumproduct_analytics_addresses(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
) -> tuple[str, ...]:
    """Return bound addresses for ``sumproduct_analytics`` output leaves."""
    leaves = resolve_leaves(graph, workbook, bindings, "sumproduct_analytics", direction="output")
    return tuple(leaf["address"] for leaf in leaves)


def spot_parity_subgraphs(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
) -> dict[str, tuple[str, ...]]:
    """Spot subgraph map including dynamically resolved sumproduct addresses."""
    subgraphs = dict(SPOT_PARITY_SUBGRAPHS)
    subgraphs["sumproduct_analytics"] = sumproduct_analytics_addresses(graph, workbook, bindings)
    return subgraphs


def spot_output_leaves(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
) -> list[OutputLeaf]:
    """Output leaves for Phase 3 spot parity series."""
    return [
        leaf
        for leaf in collect_output_leaves(graph, workbook, bindings)
        if leaf.series_id in SPOT_OUTPUT_SERIES
    ]


def modular_binding_compute_value(pkg: Any, leaf: OutputLeaf, *, ctx: Any | None = None) -> object:
    """Return ``OBS_VALUE`` from a modular ``compute_*`` function."""
    make_context = pkg.make_context
    if ctx is None:
        ctx = make_context()
    compute = getattr(pkg, f"compute_{leaf.series_id}")
    records = compute(ctx=ctx)
    record = next(
        record
        for record in records
        if all(record.get(key) == value for key, value in leaf.key.items())
    )
    return record["OBS_VALUE"]


def assert_modular_binding_matches_evaluator(
    pkg: Any,
    graph: DependencyGraph,
    leaves: list[OutputLeaf],
    eval_results: dict[str, object],
    *,
    rtol: float = RTOL,
    atol: float = ATOL,
) -> None:
    """Assert modular binding computes match evaluator for each leaf."""
    ctx = pkg.make_context()
    mismatches: list[tuple[str, object, object]] = []
    for leaf in leaves:
        if leaf.address in XFAIL_BINDING_EVAL:
            continue
        binding_value = modular_binding_compute_value(pkg, leaf, ctx=ctx)
        eval_value = eval_results.get(leaf.address)
        if not parity_values_equal(binding_value, eval_value, rtol=rtol, atol=atol):
            mismatches.append((leaf.address, eval_value, binding_value))
    if mismatches:
        lines = ["Modular binding vs evaluator mismatches:"]
        for address, eval_value, binding_value in mismatches:
            lines.append(f"- {address}: evaluator={eval_value!r} binding={binding_value!r}")
        raise AssertionError("\n".join(lines))


DOWNSTREAM_CHAIN_CASES: tuple[DownstreamChainCase, ...] = (
    DownstreamChainCase(
        name="bull_scenario_to_enterprise_value",
        setter_name="set_scenario_assumptions",
        setter_records=({"PARAMETER": "Active Scenario (1-3)", "OBS_VALUE": 3.0},),
        steps=(
            DownstreamChainStep(
                "compute_active_growth_rate",
                {"PARAMETER": "Active Growth Rate"},
                0.145,
            ),
            DownstreamChainStep(
                "compute_valuation_verdict",
                {"METRIC": "Enterprise Value"},
                pytest.approx(255.99019812632127, rel=1e-4),
            ),
        ),
    ),
    DownstreamChainCase(
        name="lookup_sku_to_panel",
        setter_name="set_lookup_sku",
        setter_records=({"FIELD": "Lookup SKU", "OBS_VALUE": "PRD-001"},),
        steps=(
            DownstreamChainStep(
                "compute_lookup_panel",
                {"FIELD": "Product Name (VLOOKUP)"},
                "Cloud Suite Pro",
            ),
            DownstreamChainStep(
                "compute_lookup_panel",
                {"FIELD": "List Price (INDEX/MATCH)"},
                1499.0,
            ),
        ),
    ),
    DownstreamChainCase(
        name="lookup_sku_to_sumproduct",
        setter_name="set_lookup_sku",
        setter_records=({"FIELD": "Lookup SKU", "OBS_VALUE": "PRD-001"},),
        steps=(
            DownstreamChainStep(
                "compute_lookup_panel",
                {"FIELD": "Product Name (VLOOKUP)"},
                "Cloud Suite Pro",
            ),
            DownstreamChainStep(
                "compute_sumproduct_analytics",
                {"FIELD": "Software Revenue Potential"},
                0.0,
            ),
        ),
        xfail=True,
        xfail_reason="SUMPRODUCT analytics returns #VALUE! after lookup SKU change",
    ),
    DownstreamChainCase(
        name="monthly_revenue_to_stats",
        setter_name="set_monthly_revenue",
        setter_records=({"MONTH": "Feb", "YEAR": 2023, "OBS_VALUE": 99.0},),
        steps=(
            DownstreamChainStep(
                "compute_mom_growth",
                {"MONTH": "Feb", "YEAR": 2023},
                pytest.approx(16.06896551724138, rel=1e-9),
            ),
            DownstreamChainStep(
                "compute_rolling_average",
                {"MONTH": "Mar", "YEAR": 2023},
                pytest.approx(37.06666666666667, rel=1e-9),
            ),
            DownstreamChainStep(
                "compute_descriptive_statistics",
                {"STATISTIC": "Mean ($M)"},
                pytest.approx(14.254166666666668, rel=1e-9),
            ),
        ),
    ),
    DownstreamChainCase(
        name="q1_sales_to_rank",
        setter_name="set_quarterly_sales_q1",
        setter_records=({"REP_ID": "R01", "OBS_VALUE": 500_000.0},),
        steps=(
            DownstreamChainStep(
                "compute_sales_pct_to_target",
                {"REP_ID": "R01"},
                pytest.approx(1.2617391304347827, rel=1e-9),
            ),
            DownstreamChainStep(
                "compute_sales_annual_total",
                {"REP_ID": "R01"},
                1_451_000.0,
            ),
            DownstreamChainStep(
                "compute_sales_rank",
                {"REP_ID": "R01"},
                3.0,
            ),
        ),
    ),
    DownstreamChainCase(
        name="product_price_to_margin",
        setter_name="set_product_list_price",
        setter_records=({"SKU": "PRD-001", "OBS_VALUE": 2000.0},),
        steps=(
            DownstreamChainStep(
                "compute_product_margin_dollars",
                {"SKU": "PRD-001"},
                1820.0,
            ),
            DownstreamChainStep(
                "compute_product_margin_pct",
                {"SKU": "PRD-001"},
                0.91,
            ),
        ),
    ),
)


def assert_downstream_propagation_chain(pkg: Any, chain_case: DownstreamChainCase) -> None:
    """Apply a setter and assert ordered downstream compute observations."""
    ctx = pkg.make_context()
    setter = getattr(pkg, chain_case.setter_name)
    setter(ctx, list(chain_case.setter_records))
    for step in chain_case.steps:
        compute = getattr(pkg, step.compute_name)
        records = compute(ctx=ctx)
        record = next(
            record
            for record in records
            if all(record.get(key) == value for key, value in step.record_key.items())
        )
        assert record["OBS_VALUE"] == step.expected_obs_value
