"""Codegen correctness tests for ``financial_model.xlsx`` binding shards.

Exercises three-way parity (Excel cache, ``FormulaEvaluator``, generated code),
per-cell accuracy, export parity, and multi-hop downstream propagation.
Known gaps are marked ``xfail`` at the cell or chain level.

The exhaustive per-cell parity suite (items 1–4) is **skipped by default**; set
``SKIP_EXHAUSTIVE_PARITY = False`` in this module to enable it locally.
"""

from __future__ import annotations

import importlib
import sys
import tempfile
from pathlib import Path

import pytest

from excel_grapher.series_bindings import load_series_bindings
from excel_grapher.series_bindings.workflow import (
    all_series_targets,
    generate_bindings_modules,
    validate_bindings_workbook,
)
from tests.integration.user_flows.bindings_accuracy import (
    BindingsAccuracyCase,
    SeriesSpotCheck,
    apply_setter,
    assert_bindings_validate,
    assert_series_spot_check,
    build_dependency_graph,
    compute_records,
    generate_bindings_namespace,
    read_workbook_cell,
)
from tests.integration.user_flows.financial_model_parity import (
    DOWNSTREAM_CHAIN_CASES,
    XFAIL_EVAL_CODEGEN,
    OutputLeaf,
    ParityLeg,
    assert_parity_equal,
    binding_compute_value,
    collect_output_leaves,
    evaluate_all_targets,
    maybe_xfail_address,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator

EXAMPLES = Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks"
WORKBOOK = EXAMPLES / "financial_model.xlsx"
BINDINGS_DIR = EXAMPLES / "financial_model.bindings"

PARITY_LEGS: tuple[ParityLeg, ...] = ("codegen_excel", "eval_excel", "eval_codegen")

# Exhaustive per-cell parity is opt-in; default CI runs validation + input spot checks only.
SKIP_EXHAUSTIVE_PARITY = True
SKIP_EXHAUSTIVE_PARITY_REASON = (
    "Exhaustive financial_model parity suite disabled by default "
    "(453 three-way + 191 export-parity tests). "
    "Set SKIP_EXHAUSTIVE_PARITY = False in this module to enable."
)

pytestmark_exhaustive = pytest.mark.skip(reason=SKIP_EXHAUSTIVE_PARITY_REASON)


def _skip_if_fixtures_missing() -> None:
    if not WORKBOOK.is_file():
        pytest.skip(f"Workbook fixture missing: {WORKBOOK}")
    if not BINDINGS_DIR.is_dir():
        pytest.skip(f"Bindings directory missing: {BINDINGS_DIR}")


def _load_output_leaves_for_collection() -> list[OutputLeaf]:
    if not WORKBOOK.is_file() or not BINDINGS_DIR.is_dir():
        return []
    bindings = load_series_bindings(BINDINGS_DIR)
    graph = build_dependency_graph(WORKBOOK, bindings, use_cached_dynamic_refs=True)
    return collect_output_leaves(graph, WORKBOOK, bindings)


def _load_targets_for_collection() -> list[str]:
    if not WORKBOOK.is_file() or not BINDINGS_DIR.is_dir():
        return []
    bindings = load_series_bindings(BINDINGS_DIR)
    return all_series_targets(bindings, workbook=WORKBOOK)


def pytest_generate_tests(metafunc: pytest.Metafunc) -> None:
    if SKIP_EXHAUSTIVE_PARITY:
        if "output_leaf" in metafunc.fixturenames and "parity_leg" in metafunc.fixturenames:
            metafunc.parametrize(
                ("output_leaf", "parity_leg"),
                [
                    pytest.param(
                        None,
                        "codegen_excel",
                        marks=pytest.mark.skip(reason=SKIP_EXHAUSTIVE_PARITY_REASON),
                    )
                ],
            )
            return
        if "binding_target" in metafunc.fixturenames:
            metafunc.parametrize(
                "binding_target",
                [
                    pytest.param(
                        None,
                        marks=pytest.mark.skip(reason=SKIP_EXHAUSTIVE_PARITY_REASON),
                    )
                ],
            )
            return
    if "output_leaf" in metafunc.fixturenames and "parity_leg" in metafunc.fixturenames:
        leaves = _load_output_leaves_for_collection()
        params = [
            pytest.param(leaf, leg, id=f"{leg}::{leaf.address}")
            for leg in PARITY_LEGS
            for leaf in leaves
        ]
        if not params:
            params = [
                pytest.param(
                    None, "codegen_excel", marks=pytest.mark.skip(reason="fixtures missing")
                )
            ]
        metafunc.parametrize(("output_leaf", "parity_leg"), params)
    if "binding_target" in metafunc.fixturenames:
        targets = _load_targets_for_collection()
        if not targets:
            metafunc.parametrize(
                "binding_target",
                [pytest.param(None, marks=pytest.mark.skip(reason="fixtures missing"))],
            )
        else:
            metafunc.parametrize("binding_target", targets, ids=targets)


@pytest.fixture(scope="module")
def workbook() -> Path:
    _skip_if_fixtures_missing()
    return WORKBOOK


@pytest.fixture(scope="module")
def bindings(workbook: Path):
    return load_series_bindings(BINDINGS_DIR)


@pytest.fixture(scope="module")
def graph(workbook: Path, bindings):
    return build_dependency_graph(workbook, bindings, use_cached_dynamic_refs=True)


@pytest.fixture(scope="module")
def namespace(workbook: Path, bindings, graph):
    return generate_bindings_namespace(graph, workbook, bindings)


@pytest.fixture(scope="module")
def parity_results(workbook: Path, bindings, graph):
    return evaluate_all_targets(graph, workbook, bindings)


@pytest.fixture(scope="module")
def eval_results(parity_results):
    return parity_results[1]


@pytest.fixture(scope="module")
def gen_cache(parity_results):
    return parity_results[2]


def test_merged_bindings_validate(workbook: Path) -> None:
    """All shards load together without schema or resolution errors."""
    case = BindingsAccuracyCase(
        name="financial_model",
        workbook=workbook,
        bindings_path=BINDINGS_DIR,
        expected_setter_count=6,
        expected_compute_count=29,
    )
    assert_bindings_validate(case)


@pytest.mark.parametrize(
    "bindings_file,expected_setters,expected_computes",
    [
        ("assumptions.bindings.yaml", 2, 1),
        ("revenue_model.bindings.yaml", 0, 11),
        ("product_lookup.bindings.yaml", 3, 3),
        ("statistical_analysis.bindings.yaml", 1, 4),
        ("dcf_valuation.bindings.yaml", 0, 10),
    ],
    ids=["assumptions", "revenue_model", "product_lookup", "statistical_analysis", "dcf_valuation"],
)
def test_shard_bindings_validate(
    workbook: Path,
    bindings_file: str,
    expected_setters: int,
    expected_computes: int,
) -> None:
    case = BindingsAccuracyCase(
        name=bindings_file.removesuffix(".bindings.yaml"),
        workbook=workbook,
        bindings_path=BINDINGS_DIR / bindings_file,
        expected_setter_count=expected_setters,
        expected_compute_count=expected_computes,
    )
    assert_bindings_validate(case)


@pytest.mark.parametrize(
    "series_id,leaf_count,sample_key,sample_value",
    [
        ("model_assumptions", 7, {"PARAMETER": "Base Revenue (Year 1)"}, 5_000_000.0),
        ("scenario_assumptions", 4, {"PARAMETER": "Selected Scenario"}, 2.0),
        ("product_prices", 8, {"PRODUCT_ID": "P001"}, 149.99),
        ("lookup_product_id", 1, {"FIELD": "Lookup Product ID"}, "P003"),
        ("monthly_sales", 12, {"MONTH": "Jan"}, 312.0),
    ],
    ids=[
        "model_assumptions",
        "scenario_assumptions",
        "product_prices",
        "lookup_product_id",
        "monthly_sales",
    ],
)
def test_input_series_spot_checks(
    workbook: Path,
    bindings,
    graph,
    series_id: str,
    leaf_count: int,
    sample_key: dict[str, str],
    sample_value: object,
) -> None:
    check = SeriesSpotCheck(
        series_id=series_id,
        leaf_count=leaf_count,
        sample_key=sample_key,
        sample_value=sample_value,
    )
    assert_series_spot_check(graph, workbook, bindings, check)


@pytestmark_exhaustive
def test_output_leaf_three_way_parity(
    workbook: Path,
    namespace: dict[str, object],
    eval_results: dict[str, object],
    gen_cache: dict[str, object],
    output_leaf: OutputLeaf | None,
    parity_leg: ParityLeg,
) -> None:
    """Per-cell parity for codegen ↔ Excel, evaluator ↔ Excel, evaluator ↔ codegen."""
    if output_leaf is None:
        pytest.skip("fixtures missing")
    maybe_xfail_address(output_leaf.address, parity_leg)

    excel_value = read_workbook_cell(workbook, output_leaf.address)
    binding_value = binding_compute_value(namespace, output_leaf)
    eval_value = eval_results.get(output_leaf.address)
    codegen_value = gen_cache.get(output_leaf.address)

    if parity_leg == "codegen_excel":
        assert_parity_equal(binding_value, excel_value, address=output_leaf.address, leg=parity_leg)
    elif parity_leg == "eval_excel":
        assert_parity_equal(eval_value, excel_value, address=output_leaf.address, leg=parity_leg)
    else:
        assert_parity_equal(eval_value, codegen_value, address=output_leaf.address, leg=parity_leg)


@pytestmark_exhaustive
def test_binding_target_evaluator_matches_codegen(
    graph,
    binding_target: str | None,
) -> None:
    """Export parity per bound target cell."""
    if binding_target is None:
        pytest.skip("fixtures missing")
    if binding_target in XFAIL_EVAL_CODEGEN:
        pytest.xfail(f"known eval_codegen gap at {binding_target}")
    assert_codegen_matches_evaluator(graph, [binding_target], rtol=1e-9, atol=1e-9)


@pytestmark_exhaustive
@pytest.mark.parametrize(
    "chain_case",
    [
        pytest.param(case, marks=pytest.mark.xfail(reason=case.xfail_reason, strict=True))
        if case.xfail
        else case
        for case in DOWNSTREAM_CHAIN_CASES
    ],
    ids=[case.name for case in DOWNSTREAM_CHAIN_CASES],
)
def test_downstream_propagation_chain(
    namespace: dict[str, object],
    chain_case,
) -> None:
    """Setter writes propagate through dependent compute functions."""
    ctx = apply_setter(namespace, chain_case.setter_name, chain_case.setter_records)
    for step in chain_case.steps:
        records = compute_records(namespace, step.compute_name, ctx=ctx)
        record = next(
            record
            for record in records
            if all(record.get(key) == value for key, value in step.record_key.items())
        )
        assert record["OBS_VALUE"] == step.expected_obs_value


@pytestmark_exhaustive
@pytest.mark.xfail(
    reason="modular export references xl_false but runtime does not export it",
    strict=True,
)
def test_product_lookup_modular_export_import(workbook: Path) -> None:
    """Modular codegen for VLOOKUP(FALSE) should import cleanly once xl_false lands."""
    bindings_path = BINDINGS_DIR / "product_lookup.bindings.yaml"
    result = validate_bindings_workbook(workbook, bindings_path)
    files = generate_bindings_modules(
        result["graph"],
        targets=result["targets"],
        bindings=result["bindings"],
        workbook=workbook,
    )
    with tempfile.TemporaryDirectory() as temp_dir:
        module_dir = Path(temp_dir) / "bindings_module"
        module_dir.mkdir()
        for filename, content in files.items():
            (module_dir / filename).write_text(content, encoding="utf-8")
        sys.path.insert(0, temp_dir)
        try:
            importlib.import_module("bindings_module")
        finally:
            sys.path.pop(0)
            sys.modules.pop("bindings_module", None)


@pytestmark_exhaustive
def test_parity_coverage_has_no_unexpected_failures(
    workbook: Path,
    bindings,
    graph,
    namespace: dict[str, object],
    eval_results: dict[str, object],
    gen_cache: dict[str, object],
) -> None:
    """Guard: newly failing cells outside known xfail sets fail the suite."""
    from tests.integration.user_flows.financial_model_parity import (
        XFAIL_CODEGEN_EXCEL,
        XFAIL_EVAL_CODEGEN,
        XFAIL_EVAL_EXCEL,
        parity_values_equal,
    )

    unexpected: list[str] = []
    for leaf in collect_output_leaves(graph, workbook, bindings):
        excel_value = read_workbook_cell(workbook, leaf.address)
        binding_value = binding_compute_value(namespace, leaf)
        eval_value = eval_results.get(leaf.address)
        codegen_value = gen_cache.get(leaf.address)

        if (
            not parity_values_equal(binding_value, excel_value)
            and leaf.address not in XFAIL_CODEGEN_EXCEL
        ):
            unexpected.append(f"codegen_excel:{leaf.address}")
        if (
            not parity_values_equal(eval_value, excel_value)
            and leaf.address not in XFAIL_EVAL_EXCEL
        ):
            unexpected.append(f"eval_excel:{leaf.address}")
        if (
            not parity_values_equal(eval_value, codegen_value)
            and leaf.address not in XFAIL_EVAL_CODEGEN
        ):
            unexpected.append(f"eval_codegen:{leaf.address}")

    assert not unexpected, "Unexpected parity failures:\n" + "\n".join(unexpected)
