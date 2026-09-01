"""Layer B — Tiny DSA inverted-tree integration canary."""

from __future__ import annotations

import inspect
from collections.abc import Callable
from typing import Annotated, Literal

import pytest

from excel_grapher.core.cell_types import Between, RealBetween
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


_COLS = ("C", "D", "E", "F", "G")
_TINY_DSA_CONSTRAINTS: dict[str, object] = {
    "Inputs!A10": Literal["Borvelia"],
    "Inputs!A11": Literal["Litellia"],
    "Inputs!A12": Literal["Aurelium"],
    "Inputs!B22": Literal[1, 2, 3],
    "Inputs!B5": Literal["Borvelia", "Litellia", "Aurelium"],
    "Engine!C5": Literal[1],
    "Engine!D5": Literal[2],
    "Engine!E5": Literal[3],
    "Engine!F5": Literal[4],
    "Engine!G5": Literal[5],
    "Inputs!B10": Annotated[float, RealBetween(0.0, 200.0)],
    "Inputs!B11": Annotated[float, RealBetween(0.0, 200.0)],
    "Inputs!B12": Annotated[float, RealBetween(0.0, 200.0)],
    "Inputs!B21": Annotated[int, Between(1, 5)],
    "Inputs!B26": Annotated[float, RealBetween(-30.0, 30.0)],
    "Inputs!C26": Annotated[float, RealBetween(-30.0, 30.0)],
    "Inputs!D26": Annotated[float, RealBetween(-30.0, 30.0)],
    **{f"Inputs!{c}16": Annotated[float, RealBetween(-10.0, 15.0)] for c in _COLS},
    **{f"Inputs!{c}17": Annotated[float, RealBetween(0.0, 20.0)] for c in _COLS},
    **{f"Inputs!{c}18": Annotated[float, RealBetween(-15.0, 15.0)] for c in _COLS},
}


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
    assert "country_profile_names" in names


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
    assert "engine_year_labels" in names
    assert "country_profile_names" in names
    assert "ctx" not in names
    assert _required(tiny_dsa_pkg.compute_output_delta) == required


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


def test_year_prefix_does_not_require_full_constant_labels(tiny_dsa_pkg) -> None:
    data = tiny_dsa_pkg.data
    source = inspect.getsource(tiny_dsa_pkg.compute_output_shocked)
    assert "require_aligned(growth_baseline, interest_baseline, primary_balance_baseline)" in source
    assert "engine_year_labels" not in source.split("horizon =", 1)[1].split("\n", 1)[0]
    result = tiny_dsa_pkg.compute_output_shocked(
        country_name=data.COUNTRY_NAME_DEFAULT,
        country_initial_debt=data.COUNTRY_INITIAL_DEBT_DEFAULT,
        growth_baseline=data.GROWTH_BASELINE_DEFAULT[:1],
        interest_baseline=data.INTEREST_BASELINE_DEFAULT[:1],
        primary_balance_baseline=data.PRIMARY_BALANCE_BASELINE_DEFAULT[:1],
        shock_year=data.SHOCK_YEAR_DEFAULT,
        shock_type=data.SHOCK_TYPE_DEFAULT,
        shock_magnitudes=data.SHOCK_MAGNITUDES_DEFAULT,
    )
    assert result == pytest.approx((_DEFAULT_SHOCKED[0],), abs=1e-9)
