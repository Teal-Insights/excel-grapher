"""Per-sheet binding cases for ``advanced_formula_workbook.xlsx``."""

from __future__ import annotations

from pathlib import Path

from tests.integration.user_flows.bindings_accuracy import (
    BindingsAccuracyCase,
    DownstreamUpdateCase,
    SeriesSpotCheck,
)

SANDBOX = Path(__file__).resolve().parents[3] / "sandbox" / "model"
WORKBOOK = SANDBOX / "advanced_formula_workbook.xlsx"
BINDINGS_DIR = SANDBOX / "advanced_formula_workbook.bindings"

MERGED_BINDINGS_CASE = BindingsAccuracyCase(
    name="advanced_formula_workbook",
    workbook=WORKBOOK,
    bindings_path=BINDINGS_DIR,
    expected_setter_count=21,
    expected_compute_count=56,
)

SHARD_BINDING_CASES: tuple[BindingsAccuracyCase, ...] = (
    BindingsAccuracyCase(
        name="assumptions",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "assumptions.bindings.yaml",
        expected_setter_count=4,
        expected_compute_count=2,
    ),
    BindingsAccuracyCase(
        name="revenue_model",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "revenue_model.bindings.yaml",
        expected_setter_count=0,
        expected_compute_count=18,
    ),
    BindingsAccuracyCase(
        name="employee_directory",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "employee_directory.bindings.yaml",
        expected_setter_count=7,
        expected_compute_count=5,
    ),
    BindingsAccuracyCase(
        name="product_lookup",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "product_lookup.bindings.yaml",
        expected_setter_count=3,
        expected_compute_count=4,
    ),
    BindingsAccuracyCase(
        name="sales_tracker",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "sales_tracker.bindings.yaml",
        expected_setter_count=5,
        expected_compute_count=5,
    ),
    BindingsAccuracyCase(
        name="statistical_analysis",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "statistical_analysis.bindings.yaml",
        expected_setter_count=1,
        expected_compute_count=7,
    ),
    BindingsAccuracyCase(
        name="dcf_valuation",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "dcf_valuation.bindings.yaml",
        expected_setter_count=0,
        expected_compute_count=14,
    ),
    BindingsAccuracyCase(
        name="formula_toolkit",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "formula_toolkit.bindings.yaml",
        expected_setter_count=0,
        expected_compute_count=1,
    ),
    BindingsAccuracyCase(
        name="function_index",
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / "function_index.bindings.yaml",
        expected_setter_count=1,
        expected_compute_count=0,
    ),
)

SERIES_SPOT_CHECKS: tuple[SeriesSpotCheck, ...] = (
    SeriesSpotCheck(
        series_id="financial_assumptions",
        direction="input",
        leaf_count=10,
        sample_key={"PARAMETER": "Revenue Growth Rate"},
        sample_value=0.082,
        unique_key_fields=("PARAMETER",),
    ),
    SeriesSpotCheck(
        series_id="active_growth_rate",
        direction="output",
        leaf_count=1,
        sample_key={"PARAMETER": "Active Growth Rate"},
        unique_key_fields=("PARAMETER",),
    ),
    SeriesSpotCheck(
        series_id="revenue_base_revenue",
        direction="output",
        leaf_count=7,
        sample_key={"FISCAL_YEAR": 2024},
        unique_key_fields=("FISCAL_YEAR",),
    ),
    SeriesSpotCheck(
        series_id="revenue_summary",
        direction="output",
        leaf_count=12,
        sample_key={"METRIC": "Total 7-Yr Revenue"},
        unique_key_fields=("METRIC",),
    ),
    SeriesSpotCheck(
        series_id="employee_salary",
        direction="input",
        leaf_count=20,
        sample_key={"EMP_ID": "E001"},
        sample_value=112_000.0,
        unique_key_fields=("EMP_ID",),
    ),
    SeriesSpotCheck(
        series_id="employee_tenure",
        direction="output",
        leaf_count=20,
        sample_key={"EMP_ID": "E001"},
        unique_key_fields=("EMP_ID",),
    ),
    SeriesSpotCheck(
        series_id="lookup_sku",
        direction="input",
        leaf_count=1,
        sample_key={"FIELD": "Lookup SKU"},
        sample_value="PRD-008",
        unique_key_fields=("FIELD",),
    ),
    SeriesSpotCheck(
        series_id="lookup_panel",
        direction="output",
        leaf_count=14,
        sample_key={"FIELD": "Product Name (VLOOKUP)"},
        unique_key_fields=("FIELD",),
    ),
    SeriesSpotCheck(
        series_id="quarterly_sales_q1",
        direction="input",
        leaf_count=12,
        sample_key={"REP_ID": "R01"},
        sample_value=284_000.0,
        unique_key_fields=("REP_ID",),
    ),
    SeriesSpotCheck(
        series_id="sales_rank",
        direction="output",
        leaf_count=12,
        sample_key={"REP_ID": "R01"},
        unique_key_fields=("REP_ID",),
    ),
    SeriesSpotCheck(
        series_id="monthly_revenue",
        direction="input",
        leaf_count=24,
        sample_key={"MONTH": "Jan", "YEAR": 2023},
        sample_value=5.8,
        unique_key_fields=("MONTH", "YEAR"),
    ),
    SeriesSpotCheck(
        series_id="normdist_sigma_band",
        direction="output",
        leaf_count=10,
        sample_key={"MONTH": "Mar", "YEAR": 2024},
        unique_key_fields=("MONTH", "YEAR"),
    ),
    SeriesSpotCheck(
        series_id="dcf_free_cash_flow",
        direction="output",
        leaf_count=6,
        sample_key={"FISCAL_YEAR": 2025},
        unique_key_fields=("FISCAL_YEAR",),
    ),
    SeriesSpotCheck(
        series_id="toolkit_demos",
        direction="output",
        leaf_count=26,
        sample_key={"DEMO": "LEFT"},
        unique_key_fields=("DEMO",),
    ),
    SeriesSpotCheck(
        series_id="function_names",
        direction="input",
        leaf_count=44,
        sample_key={"FUNCTION": "IF"},
        sample_value="IF",
        unique_key_fields=("FUNCTION",),
    ),
)

DOWNSTREAM_UPDATE_CASES: tuple[DownstreamUpdateCase, ...] = (
    DownstreamUpdateCase(
        setter_name="set_scenario_assumptions",
        setter_records=({"PARAMETER": "Active Scenario (1-3)", "OBS_VALUE": 3.0},),
        compute_name="compute_active_growth_rate",
        record_key={"PARAMETER": "Active Growth Rate"},
        expected_obs_value=0.145,
    ),
    DownstreamUpdateCase(
        setter_name="set_lookup_sku",
        setter_records=({"FIELD": "Lookup SKU", "OBS_VALUE": "PRD-001"},),
        compute_name="compute_lookup_panel",
        record_key={"FIELD": "List Price (INDEX/MATCH)"},
        expected_obs_value=1499.0,
    ),
    DownstreamUpdateCase(
        setter_name="set_monthly_revenue",
        setter_records=({"MONTH": "Feb", "YEAR": 2023, "OBS_VALUE": 99.0},),
        compute_name="compute_mom_growth",
        record_key={"MONTH": "Feb", "YEAR": 2023},
        expected_obs_value=16.06896551724138,
    ),
    DownstreamUpdateCase(
        setter_name="set_quarterly_sales_q1",
        setter_records=({"REP_ID": "R01", "OBS_VALUE": 500_000.0},),
        compute_name="compute_sales_pct_to_target",
        record_key={"REP_ID": "R01"},
        expected_obs_value=1.2617391304347827,
    ),
)
