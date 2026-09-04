"""Layer B — Tiny DSA inverted-tree integration canary."""

from __future__ import annotations

import inspect
from collections.abc import Callable

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings.load import load_series_bindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.paths import INVERTED_TREE_TINY_DSA
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    load_package,
    required_param_names,
)
from tests.unit.exporter.inverted_tree.local_corpus import load_constraints_module

_WORKBOOK = INVERTED_TREE_TINY_DSA / "tiny-dsa.xlsx"
_BINDINGS_DIR = INVERTED_TREE_TINY_DSA / "bindings"

_DEFAULT_BASELINE = (
    61.28985507246378,
    62.0859413288525,
    62.38587341256677,
    62.18725444354536,
    61.48767596259631,
)
_DEFAULT_SHOCKED = (
    61.28985507246378,
    63.29945741415009,
    64.85855735045921,
    65.95605876303210,
    66.58059223010186,
)


def _required(function: Callable[..., object]) -> tuple[str, ...]:
    return required_param_names(function)


_tiny_dsa_constraints_mod = load_constraints_module(INVERTED_TREE_TINY_DSA / "constraints.py")
assert _tiny_dsa_constraints_mod is not None
_TINY_DSA_CONSTRAINTS = _tiny_dsa_constraints_mod.CONSTRAINTS


def _tiny_dsa_graph():
    bindings = load_series_bindings(_BINDINGS_DIR)
    targets = all_series_targets(bindings, workbook=_WORKBOOK)
    graph = create_dependency_graph(
        _WORKBOOK,
        targets,
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_constraints(_TINY_DSA_CONSTRAINTS, {}),
    )
    return bindings, targets, graph


@pytest.fixture(scope="module")
def tiny_dsa_pkg(tmp_path_factory: pytest.TempPathFactory):
    bindings, targets, graph = _tiny_dsa_graph()
    with CodeGenerator(graph) as gen:
        modules = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=_WORKBOOK,
            paradigm="inverted_tree",
        )
    tmp_path = tmp_path_factory.mktemp("tiny_dsa_inverted")
    return load_package(modules, tmp_path, name="tiny_dsa_inverted")


def test_helper_inventory_matches_bound_formula_series(tiny_dsa_pkg) -> None:
    internals = tiny_dsa_pkg.internals
    for name in (
        "initial_debt_resolved",
        "shock_active",
        "shocked_growth",
        "baseline_path_internal",
        "output_baseline",
        "shocked_path_internal",
        "output_shocked",
        "output_delta",
        "shock_magnitude_resolved",
    ):
        assert callable(getattr(internals, name))
    source = inspect.getsource(internals)
    assert "def cell_" not in source
    assert "def make_context" not in inspect.getsource(tiny_dsa_pkg.api)
    assert "def set_" not in inspect.getsource(tiny_dsa_pkg.api)
    api_src = inspect.getsource(tiny_dsa_pkg.api)
    assert api_src.count("internals.shocked_path_internal(") == 1


def test_baseline_leaf_closure_excludes_shock_args(tiny_dsa_pkg) -> None:
    required = _required(tiny_dsa_pkg.compute_output_baseline)
    names = all_param_names(tiny_dsa_pkg.compute_output_baseline)
    assert required == (
        "country_name",
        "country_initial_debt",
        "growth_baseline",
        "interest_baseline",
        "primary_balance_baseline",
    )
    assert "shock_year" not in names
    assert "shock_type" not in names
    assert "shock_magnitudes" not in names
    assert "engine_year_labels" not in names
    assert "ctx" not in names
    assert "country_profile_names" not in names
    assert tiny_dsa_pkg.compute_output_baseline.__constants__ == ("country_profile_names",)


def test_shocked_leaf_closure_includes_shock_args(tiny_dsa_pkg) -> None:
    required = _required(tiny_dsa_pkg.compute_output_shocked)
    names = all_param_names(tiny_dsa_pkg.compute_output_shocked)
    assert required == (
        "country_name",
        "country_initial_debt",
        "growth_baseline",
        "interest_baseline",
        "primary_balance_baseline",
        "shock_year",
        "shock_type",
        "shock_magnitudes",
    )
    assert "engine_year_labels" not in names
    assert "country_profile_names" not in names
    assert "ctx" not in names
    assert _required(tiny_dsa_pkg.compute_output_delta) == required
    assert tiny_dsa_pkg.compute_output_shocked.__constants__ == (
        "country_profile_names",
        "engine_year_labels",
    )
    assert tiny_dsa_pkg.compute_output_delta.__constants__ == (
        "country_profile_names",
        "engine_year_labels",
    )


def test_shocked_path_internal_first_level_deps(tiny_dsa_pkg) -> None:
    names = all_param_names(tiny_dsa_pkg.internals.shocked_path_internal)
    assert "shocked_growth" in names
    assert "shocked_interest" in names
    assert "shocked_primary_balance" in names
    assert "shock_type" not in names
    assert "shock_magnitudes" not in names
    assert "ctx" not in names
    assert any("initial_debt" in name for name in names)


def test_shock_active_params(tiny_dsa_pkg) -> None:
    assert _required(tiny_dsa_pkg.internals.shock_active) == (
        "engine_year_labels",
        "shock_year",
    )


def test_default_borvelia_numeric_parity(tiny_dsa_pkg) -> None:
    data = tiny_dsa_pkg.data
    baseline = tiny_dsa_pkg.compute_output_baseline(
        country_name=data.COUNTRY_NAME_DEFAULT,
        country_initial_debt=data.COUNTRY_INITIAL_DEBT_DEFAULT,
        growth_baseline=data.GROWTH_BASELINE_DEFAULT,
        interest_baseline=data.INTEREST_BASELINE_DEFAULT,
        primary_balance_baseline=data.PRIMARY_BALANCE_BASELINE_DEFAULT,
    )
    shocked = tiny_dsa_pkg.compute_output_shocked(
        country_name=data.COUNTRY_NAME_DEFAULT,
        country_initial_debt=data.COUNTRY_INITIAL_DEBT_DEFAULT,
        growth_baseline=data.GROWTH_BASELINE_DEFAULT,
        interest_baseline=data.INTEREST_BASELINE_DEFAULT,
        primary_balance_baseline=data.PRIMARY_BALANCE_BASELINE_DEFAULT,
        shock_year=data.SHOCK_YEAR_DEFAULT,
        shock_type=data.SHOCK_TYPE_DEFAULT,
        shock_magnitudes=data.SHOCK_MAGNITUDES_DEFAULT,
    )
    assert baseline == pytest.approx(_DEFAULT_BASELINE, abs=1e-9)
    assert shocked == pytest.approx(_DEFAULT_SHOCKED, abs=1e-9)


def test_formula_evaluator_parity_still_holds() -> None:
    _bindings, _targets, graph = _tiny_dsa_graph()
    evaluator = FormulaEvaluator(graph)
    outputs = evaluator.evaluate(
        [f"Outputs!{col}12" for col in "BCDEF"] + [f"Outputs!{col}13" for col in "BCDEF"]
    )
    baseline = tuple(outputs[f"Outputs!{col}12"] for col in "BCDEF")
    shocked = tuple(outputs[f"Outputs!{col}13"] for col in "BCDEF")
    assert baseline == pytest.approx(_DEFAULT_BASELINE, abs=1e-9)
    assert shocked == pytest.approx(_DEFAULT_SHOCKED, abs=1e-9)


def test_time_period_domain_matches_output_header(tiny_dsa_pkg) -> None:
    from fastpyxl import load_workbook

    book = load_workbook(_WORKBOOK, data_only=True)
    header = tuple(book["Outputs"][f"{col}11"].value for col in "BCDEF")
    assert header == tiny_dsa_pkg.data.TIME_PERIOD_DOMAIN
    data = tiny_dsa_pkg.data
    computes = (
        tiny_dsa_pkg.compute_output_baseline,
        tiny_dsa_pkg.compute_output_shocked,
        tiny_dsa_pkg.compute_output_delta,
    )
    kwargs = {
        "country_name": data.COUNTRY_NAME_DEFAULT,
        "country_initial_debt": data.COUNTRY_INITIAL_DEBT_DEFAULT,
        "growth_baseline": data.GROWTH_BASELINE_DEFAULT,
        "interest_baseline": data.INTEREST_BASELINE_DEFAULT,
        "primary_balance_baseline": data.PRIMARY_BALANCE_BASELINE_DEFAULT,
    }
    shock = {
        "shock_year": data.SHOCK_YEAR_DEFAULT,
        "shock_type": data.SHOCK_TYPE_DEFAULT,
        "shock_magnitudes": data.SHOCK_MAGNITUDES_DEFAULT,
    }
    for compute in computes:
        assert compute.__key__ == ("TIME_PERIOD",)
        assert compute.__domain__ == header
        args = kwargs if compute is tiny_dsa_pkg.compute_output_baseline else {**kwargs, **shock}
        result = compute(**args)
        assert len(compute.__domain__) == len(result)
        last_year = header[-1]
        assert result[data.TIME_PERIOD_DOMAIN.index(last_year)] == result[-1]


def test_internals_helpers_publish_key_domains(tiny_dsa_pkg) -> None:
    for name in (
        "output_baseline",
        "output_shocked",
        "output_delta",
        "baseline_path_internal",
        "shocked_path_internal",
        "shock_active",
    ):
        helper = getattr(tiny_dsa_pkg.internals, name)
        assert helper.__key__ == ("TIME_PERIOD",)
        assert len(helper.__domain__) > 0


def test_public_computes_require_catalog_order_arrays(tiny_dsa_pkg) -> None:
    source = inspect.getsource(tiny_dsa_pkg.api)
    assert "trim(" not in source
    assert "take(" not in source
    data = tiny_dsa_pkg.data
    with pytest.raises(ValueError, match="expected length"):
        tiny_dsa_pkg.compute_output_shocked(
            country_name=data.COUNTRY_NAME_DEFAULT,
            country_initial_debt=data.COUNTRY_INITIAL_DEBT_DEFAULT,
            growth_baseline=data.GROWTH_BASELINE_DEFAULT[:1],
            interest_baseline=data.INTEREST_BASELINE_DEFAULT[:1],
            primary_balance_baseline=data.PRIMARY_BALANCE_BASELINE_DEFAULT[:1],
            shock_year=data.SHOCK_YEAR_DEFAULT,
            shock_type=data.SHOCK_TYPE_DEFAULT,
            shock_magnitudes=data.SHOCK_MAGNITUDES_DEFAULT,
        )
