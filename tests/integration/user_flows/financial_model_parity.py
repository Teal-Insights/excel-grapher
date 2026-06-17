"""Parity helpers for ``financial_model.xlsx`` binding accuracy tests."""

from __future__ import annotations

from dataclasses import dataclass
from math import isfinite
from typing import Any, Literal, cast

import pytest

from excel_grapher import XlError
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.grapher import DependencyGraph
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import all_series_targets, compute_names
from tests.integration.user_flows.bindings_accuracy import (
    compute_records,
    resolve_leaves,
)
from tests.integration.utils.parity_harness import exec_generated_code_with_cache

ParityLeg = Literal["codegen_excel", "eval_excel", "eval_codegen"]

RTOL = 1e-6
ATOL = 1e-6


@dataclass(frozen=True)
class OutputLeaf:
    """One bound output cell with its series id and record key."""

    series_id: str
    address: str
    key: dict[str, object]


@dataclass(frozen=True)
class DownstreamChainStep:
    """Expected observation after one or more setter writes."""

    compute_name: str
    record_key: dict[str, object]
    expected_obs_value: object


@dataclass(frozen=True)
class DownstreamChainCase:
    """Multi-hop setter chain with ordered downstream expectations."""

    name: str
    setter_name: str
    setter_records: tuple[dict[str, object], ...]
    steps: tuple[DownstreamChainStep, ...]
    xfail: bool = False
    xfail_reason: str = ""


# Cells where generated binding computes disagree with cached Excel values.
XFAIL_CODEGEN_EXCEL: frozenset[str] = frozenset(
    {
        "'Statistical Analysis'!F6",
        "'Statistical Analysis'!F16",
        "'Statistical Analysis'!F17",
        "'Statistical Analysis'!F18",
        "'Statistical Analysis'!F19",
        "'Statistical Analysis'!F20",
        "'Statistical Analysis'!F21",
        "'Statistical Analysis'!F22",
        "'Statistical Analysis'!F23",
        "'Statistical Analysis'!F24",
        "'Statistical Analysis'!F25",
        "'Statistical Analysis'!G16",
        "'Statistical Analysis'!G17",
        "'Statistical Analysis'!G18",
        "'Statistical Analysis'!G19",
        "'Statistical Analysis'!G20",
        "'Statistical Analysis'!G21",
        "'Statistical Analysis'!G22",
        "'Statistical Analysis'!G23",
        "'Statistical Analysis'!G24",
        "'Statistical Analysis'!G25",
        "'Statistical Analysis'!H16",
        "'Statistical Analysis'!H17",
        "'Statistical Analysis'!H18",
        "'Statistical Analysis'!H19",
        "'Statistical Analysis'!H20",
        "'Statistical Analysis'!H21",
        "'Statistical Analysis'!H22",
        "'Statistical Analysis'!H23",
        "'Statistical Analysis'!H24",
        "'Statistical Analysis'!H25",
        "'Product Lookup'!I14",
        "'Product Lookup'!I15",
        "'Product Lookup'!I16",
        "'Product Lookup'!I18",
        "'Product Lookup'!I19",
        "'DCF Valuation'!B24",
    }
)

# Evaluator disagreements with cached Excel (includes TEXT/INDEX on revenue summary).
XFAIL_EVAL_EXCEL: frozenset[str] = XFAIL_CODEGEN_EXCEL | frozenset({"'Revenue Model'!B22"})

# Export drift: standalone generated code differs from FormulaEvaluator.
XFAIL_EVAL_CODEGEN: frozenset[str] = frozenset(
    {
        "'Statistical Analysis'!H16",
        "'Statistical Analysis'!H17",
        "'Statistical Analysis'!H18",
        "'Statistical Analysis'!H19",
        "'Statistical Analysis'!H20",
        "'Statistical Analysis'!H21",
        "'Statistical Analysis'!H22",
        "'Statistical Analysis'!H23",
        "'Statistical Analysis'!H24",
        "'Statistical Analysis'!H25",
        "'Revenue Model'!B22",
        "'Product Lookup'!I18",
    }
)


DOWNSTREAM_CHAIN_CASES: tuple[DownstreamChainCase, ...] = (
    DownstreamChainCase(
        name="bull_scenario_to_enterprise_value",
        setter_name="set_scenario_assumptions",
        setter_records=({"PARAMETER": "Selected Scenario", "OBS_VALUE": 3},),
        steps=(
            DownstreamChainStep(
                "compute_active_growth_rate",
                {"PARAMETER": "Active Growth Rate"},
                0.15,
            ),
            DownstreamChainStep(
                "compute_revenue_revenue",
                {"FISCAL_YEAR": 2029},
                pytest.approx(8745.03125, rel=1e-4),
            ),
            DownstreamChainStep(
                "compute_dcf_revenue",
                {"PROJECTION_YEAR": "Year 4"},
                pytest.approx(7604.375, rel=1e-4),
            ),
            DownstreamChainStep(
                "compute_valuation_summary_numeric",
                {"METRIC": "Enterprise Value"},
                pytest.approx(16578.137678437255, rel=1e-4),
            ),
        ),
    ),
    DownstreamChainCase(
        name="base_revenue_to_projection",
        setter_name="set_model_assumptions",
        setter_records=({"PARAMETER": "Base Revenue (Year 1)", "OBS_VALUE": 6_000_000},),
        steps=(
            DownstreamChainStep(
                "compute_revenue_revenue",
                {"FISCAL_YEAR": 2025},
                6000.0,
            ),
        ),
    ),
    DownstreamChainCase(
        name="product_price_to_revenue",
        setter_name="set_product_prices",
        setter_records=({"PRODUCT_ID": "P001", "OBS_VALUE": 200.0},),
        steps=(
            DownstreamChainStep(
                "compute_product_revenue",
                {"PRODUCT_ID": "P001"},
                640_000.0,
            ),
        ),
    ),
    DownstreamChainCase(
        name="invalid_lookup_id",
        setter_name="set_lookup_product_id",
        setter_records=({"FIELD": "Lookup Product ID", "OBS_VALUE": "P999"},),
        steps=(
            DownstreamChainStep(
                "compute_lookup_results",
                {"FIELD": "Product Name (VLOOKUP)"},
                "Not Found",
            ),
        ),
    ),
    DownstreamChainCase(
        name="monthly_sales_to_mean",
        setter_name="set_monthly_sales",
        setter_records=({"MONTH": "Jan", "OBS_VALUE": 400.0},),
        steps=(
            DownstreamChainStep(
                "compute_descriptive_statistics",
                {"STATISTIC": "Mean"},
                462.5,
            ),
        ),
    ),
    DownstreamChainCase(
        name="monthly_sales_to_std_dev",
        setter_name="set_monthly_sales",
        setter_records=({"MONTH": "Jan", "OBS_VALUE": 400.0},),
        steps=(
            DownstreamChainStep(
                "compute_descriptive_statistics",
                {"STATISTIC": "Std Dev"},
                pytest.approx(102.926214132045, rel=1e-4),
            ),
        ),
        xfail=True,
        xfail_reason="SUMPRODUCT variance formula returns #N/A in evaluator",
    ),
)


def normalize_parity_value(value: object) -> object:
    """Normalize values for cross-artifact comparison."""
    if isinstance(value, XlError):
        return str(value)
    if isinstance(value, str) and value.startswith("#"):
        return value
    return value


def parity_values_equal(
    actual: object,
    expected: object,
    *,
    rtol: float = RTOL,
    atol: float = ATOL,
) -> bool:
    """Return whether two parity values match."""
    actual_norm = normalize_parity_value(actual)
    expected_norm = normalize_parity_value(expected)
    if actual_norm == expected_norm:
        return True
    if isinstance(actual_norm, bool) or isinstance(expected_norm, bool):
        return False
    if isinstance(actual_norm, (int, float)) and isinstance(expected_norm, (int, float)):
        if not isfinite(float(actual_norm)) or not isfinite(float(expected_norm)):
            return False
        af = float(actual_norm)
        ef = float(expected_norm)
        return abs(af - ef) <= max(atol, rtol * max(abs(af), abs(ef), 1.0))
    return False


def collect_output_leaves(
    graph: DependencyGraph,
    workbook: Any,
    bindings: WorkbookSeriesBindings,
) -> list[OutputLeaf]:
    """Collect every output leaf across all declared compute series."""
    leaves: list[OutputLeaf] = []
    for compute_name in compute_names(bindings):
        series_id = compute_name.removeprefix("compute_")
        for leaf in resolve_leaves(graph, workbook, bindings, series_id, direction="output"):
            leaves.append(
                OutputLeaf(
                    series_id=series_id,
                    address=leaf["address"],
                    key=dict(leaf["key"]),
                )
            )
    return leaves


def binding_compute_value(
    namespace: dict[str, object],
    leaf: OutputLeaf,
    *,
    ctx: object | None = None,
) -> object:
    """Return ``OBS_VALUE`` from a generated binding compute function."""
    records = compute_records(namespace, f"compute_{leaf.series_id}", ctx=ctx)
    record = next(
        record
        for record in records
        if all(record.get(key) == value for key, value in leaf.key.items())
    )
    return record["OBS_VALUE"]


def assert_parity_equal(actual: object, expected: object, *, address: str, leg: ParityLeg) -> None:
    """Assert two parity values match with a readable failure."""
    if parity_values_equal(actual, expected):
        return
    raise AssertionError(f"{leg} mismatch at {address}: expected={expected!r} actual={actual!r}")


def xfail_set_for_leg(leg: ParityLeg) -> frozenset[str]:
    """Return known failing addresses for a parity leg."""
    if leg == "codegen_excel":
        return XFAIL_CODEGEN_EXCEL
    if leg == "eval_excel":
        return XFAIL_EVAL_EXCEL
    return XFAIL_EVAL_CODEGEN


def maybe_xfail_address(address: str, leg: ParityLeg) -> None:
    """Mark the current test xfail when ``address`` is a known gap for ``leg``."""
    if address in xfail_set_for_leg(leg):
        pytest.xfail(f"known {leg.replace('_', ' ')} gap at {address}")


def evaluate_all_targets(
    graph: DependencyGraph,
    workbook: Any,
    bindings: WorkbookSeriesBindings,
) -> tuple[list[str], dict[str, object], dict[str, object]]:
    """Evaluate all binding targets via evaluator and standalone codegen."""
    targets = all_series_targets(bindings, workbook=workbook)
    with FormulaEvaluator(graph) as evaluator:
        eval_results = cast(dict[str, object], evaluator.evaluate(targets))
    gen_cache, _, _ = exec_generated_code_with_cache(graph, targets)
    return targets, eval_results, gen_cache
