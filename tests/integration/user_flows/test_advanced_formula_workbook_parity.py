"""Output parity tests for ``advanced_formula_workbook.xlsx`` binding shards.

Default CI runs targeted spot subgraphs (evaluator ↔ codegen, modular binding ↔
evaluator, and input-leaf evaluator ↔ workbook cache). The exhaustive per-cell
three-way suite is opt-in; set ``SKIP_EXHAUSTIVE_PARITY = False`` to enable it.
"""

from __future__ import annotations

import importlib
import sys
import tempfile
from pathlib import Path
from typing import cast

import pytest

from excel_grapher import FormulaEvaluator
from excel_grapher.grapher import DependencyGraph
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import (
    BindingsCheckResult,
    run_binding_checks,
    validate_bindings_workbook,
)
from tests.integration.user_flows.advanced_formula_workbook_parity import (
    ATOL,
    BINDINGS_DIR,
    DOWNSTREAM_CHAIN_CASES,
    INPUT_SPOT_ADDRESSES,
    RTOL,
    SPOT_PARITY_SUBGRAPHS,
    WORKBOOK,
    XFAIL_BINDING_EVAL,
    XFAIL_CODEGEN_EXCEL,
    XFAIL_EVAL_CODEGEN,
    XFAIL_EVAL_EXCEL,
    assert_downstream_propagation_chain,
    assert_modular_binding_matches_evaluator,
    maybe_xfail_address,
    modular_binding_compute_value,
    spot_output_leaves,
    spot_parity_subgraphs,
)
from tests.integration.user_flows.bindings_accuracy import read_workbook_cell
from tests.integration.user_flows.financial_model_parity import (
    OutputLeaf,
    ParityLeg,
    assert_parity_equal,
    binding_compute_value,
    collect_output_leaves,
    evaluate_all_targets,
    parity_values_equal,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.utils.excel_workbook_parity import assert_workbook_parity

SKIP_EXHAUSTIVE_PARITY = True
SKIP_EXHAUSTIVE_PARITY_REASON = (
    "Exhaustive advanced_formula_workbook parity suite disabled by default. "
    "Set SKIP_EXHAUSTIVE_PARITY = False in this module to enable."
)

PARITY_LEGS: tuple[ParityLeg, ...] = ("codegen_excel", "eval_excel", "eval_codegen")
pytestmark_exhaustive = pytest.mark.skip(reason=SKIP_EXHAUSTIVE_PARITY_REASON)


def _skip_if_fixtures_missing() -> None:
    if not WORKBOOK.is_file() or not BINDINGS_DIR.is_dir():
        pytest.skip("sandbox advanced_formula_workbook fixtures missing")


def _load_output_leaves_for_collection() -> list[OutputLeaf]:
    if not WORKBOOK.is_file() or not BINDINGS_DIR.is_dir():
        return []
    result = validate_bindings_workbook(WORKBOOK, BINDINGS_DIR)
    return collect_output_leaves(result["graph"], WORKBOOK, result["bindings"])


def pytest_generate_tests(metafunc: pytest.Metafunc) -> None:
    if SKIP_EXHAUSTIVE_PARITY:
        return
    if {"output_leaf", "parity_leg"} <= set(metafunc.fixturenames):
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


@pytest.fixture(scope="module")
def sandbox_validation() -> BindingsCheckResult:
    _skip_if_fixtures_missing()
    result = validate_bindings_workbook(WORKBOOK, BINDINGS_DIR)
    assert result["report"]["ok"], result["report"]["issues"]
    return result


@pytest.fixture(scope="module")
def workbook(sandbox_validation: BindingsCheckResult) -> Path:
    return WORKBOOK


@pytest.fixture(scope="module")
def bindings(sandbox_validation: BindingsCheckResult) -> WorkbookSeriesBindings:
    return sandbox_validation["bindings"]


@pytest.fixture(scope="module")
def graph(sandbox_validation: BindingsCheckResult) -> DependencyGraph:
    return sandbox_validation["graph"]


@pytest.fixture(scope="module")
def spot_subgraphs(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
) -> dict[str, tuple[str, ...]]:
    return spot_parity_subgraphs(graph, workbook, bindings)


@pytest.fixture(scope="module")
def modular_package():
    _skip_if_fixtures_missing()
    with tempfile.TemporaryDirectory() as tmp:
        module_dir = Path(tmp) / "advanced_formula_workbook"
        run_binding_checks(
            WORKBOOK,
            BINDINGS_DIR,
            module_dir=module_dir,
            package_name="advanced_formula_workbook",
            smoke_test=False,
        )
        sys.path.insert(0, str(module_dir.parent))
        try:
            yield importlib.import_module("advanced_formula_workbook")
        finally:
            sys.path.remove(str(module_dir.parent))
            for name in list(sys.modules):
                if name == "advanced_formula_workbook" or name.startswith(
                    "advanced_formula_workbook."
                ):
                    sys.modules.pop(name, None)


@pytest.fixture(scope="module")
def spot_eval_results(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
) -> dict[str, object]:
    addresses = [leaf.address for leaf in spot_output_leaves(graph, workbook, bindings)]
    with FormulaEvaluator(graph) as evaluator:
        return cast(dict[str, object], evaluator.evaluate(addresses))


@pytest.mark.parametrize(
    "subgraph_name",
    list(SPOT_PARITY_SUBGRAPHS) + ["sumproduct_analytics"],
)
def test_spot_subgraph_eval_codegen_parity(
    graph: DependencyGraph,
    spot_subgraphs: dict[str, tuple[str, ...]],
    subgraph_name: str,
) -> None:
    """Evaluator and standalone codegen agree on each spot subgraph."""
    addresses = list(spot_subgraphs[subgraph_name])
    passing = [address for address in addresses if address not in XFAIL_EVAL_CODEGEN]
    assert passing, f"no passing addresses left in subgraph {subgraph_name!r}"
    assert_codegen_matches_evaluator(graph, passing, rtol=RTOL, atol=ATOL)


@pytest.mark.parametrize("address", sorted(XFAIL_EVAL_CODEGEN))
def test_spot_subgraph_eval_codegen_known_gaps(address: str, graph: DependencyGraph) -> None:
    """Document evaluator ↔ codegen gaps outside the default passing set."""
    with pytest.raises(AssertionError, match="Parity mismatch"):
        assert_codegen_matches_evaluator(graph, [address], rtol=RTOL, atol=ATOL)


def test_spot_output_series_modular_binding_matches_evaluator(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
    modular_package: object,
    spot_eval_results: dict[str, object],
) -> None:
    """Modular ``compute_*`` exports match the evaluator on spot output series."""
    leaves = spot_output_leaves(graph, workbook, bindings)
    assert_modular_binding_matches_evaluator(
        modular_package,
        graph,
        leaves,
        spot_eval_results,
        rtol=RTOL,
        atol=ATOL,
    )


def test_employee_tenure_spot_cells_are_positive_years(
    graph: DependencyGraph,
    spot_subgraphs: dict[str, tuple[str, ...]],
) -> None:
    """``TODAY``-based tenure returns positive year counts (shape check; cache may be stale)."""
    addresses = list(spot_subgraphs["employee_tenure"])
    with FormulaEvaluator(graph) as evaluator:
        for address in addresses:
            value = evaluator.evaluate(address)
            assert isinstance(value, float)
            assert value > 0.0


@pytest.mark.parametrize("address", INPUT_SPOT_ADDRESSES)
def test_spot_input_leaf_evaluator_matches_workbook_cache(
    graph: DependencyGraph,
    address: str,
) -> None:
    """Editable input leaves with cached values match the evaluator."""
    assert_workbook_parity(graph, [address], rtol=RTOL, atol=ATOL)


@pytest.mark.parametrize("address", sorted(XFAIL_BINDING_EVAL))
def test_spot_binding_eval_known_gaps(
    address: str,
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
    modular_package: object,
) -> None:
    """Document modular binding ↔ evaluator gaps on partial-overlap or formula edges."""
    leaves = [
        leaf for leaf in spot_output_leaves(graph, workbook, bindings) if leaf.address == address
    ]
    assert len(leaves) == 1
    with FormulaEvaluator(graph) as evaluator:
        eval_value = evaluator.evaluate(address)
    binding_value = modular_binding_compute_value(modular_package, leaves[0])
    assert not parity_values_equal(eval_value, binding_value, rtol=RTOL, atol=ATOL)


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
    modular_package: object,
    chain_case,
) -> None:
    """Setter writes propagate through multi-hop dependent compute functions."""
    assert_downstream_propagation_chain(modular_package, chain_case)


@pytest.fixture(scope="module")
def namespace(workbook: Path, bindings: WorkbookSeriesBindings, graph: DependencyGraph):
    from tests.integration.user_flows.bindings_accuracy import generate_bindings_namespace

    return generate_bindings_namespace(graph, workbook, bindings)


@pytest.fixture(scope="module")
def parity_results(workbook: Path, bindings: WorkbookSeriesBindings, graph: DependencyGraph):
    return evaluate_all_targets(graph, workbook, bindings)


@pytest.fixture(scope="module")
def eval_results(parity_results):
    return parity_results[1]


@pytest.fixture(scope="module")
def gen_cache(parity_results):
    return parity_results[2]


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
def test_parity_coverage_has_no_unexpected_failures(
    workbook: Path,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    namespace: dict[str, object],
    eval_results: dict[str, object],
    gen_cache: dict[str, object],
) -> None:
    """Guard: newly failing cells outside known xfail sets fail the suite."""
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
