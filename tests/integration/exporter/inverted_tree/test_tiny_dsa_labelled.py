"""Tiny DSA canary for runtime TIME_PERIOD labels (#841)."""

from __future__ import annotations

from typing import Any

import pytest

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings.load import load_series_bindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.paths import INVERTED_TREE_TINY_DSA_LABELLED
from tests.unit.exporter.inverted_tree.helpers import load_package, required_param_names
from tests.unit.exporter.inverted_tree.local_corpus import load_constraints_module

_WORKBOOK = INVERTED_TREE_TINY_DSA_LABELLED / "tiny-dsa-labelled.xlsx"
_BINDINGS_DIR = INVERTED_TREE_TINY_DSA_LABELLED / "bindings"

_DEFAULT_BASELINE = (
    61.28985507246378,
    62.0859413288525,
    62.38587341256677,
    62.18725444354536,
    61.48767596259631,
)

_constraints_mod = load_constraints_module(INVERTED_TREE_TINY_DSA_LABELLED / "constraints.py")
assert _constraints_mod is not None
_CONSTRAINTS = _constraints_mod.CONSTRAINTS


def _labelled_graph():
    bindings = load_series_bindings(_BINDINGS_DIR)
    targets = all_series_targets(bindings, workbook=_WORKBOOK)
    graph = create_dependency_graph(
        _WORKBOOK,
        targets,
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_constraints(_CONSTRAINTS, {}),
    )
    return bindings, graph


@pytest.fixture(scope="module")
def labelled_pkg(tmp_path_factory: pytest.TempPathFactory):
    bindings, graph = _labelled_graph()
    with CodeGenerator(graph) as gen:
        modules = gen.generate_modules(series_bindings=bindings, bindings_workbook=_WORKBOOK)
    tmp_path = tmp_path_factory.mktemp("tiny_dsa_labelled")
    return load_package(modules, tmp_path, name="tiny_dsa_labelled")


def _rekey(tensor: Any, years: tuple[int, ...]) -> Any:
    axis = tensor.domain.axes[0]
    shifted = type(axis)(axis.name, years, axis.key_type)
    values = [tensor[coord] for coord in tensor.domain]
    return type(tensor)(
        type(tensor.domain).product(shifted),
        tuple(values),
        schema=tensor.schema,
        cells=tensor.cells,
    )


def _baseline_kwargs(pkg, *, years: tuple[int, ...] | None = None) -> dict[str, object]:
    data = pkg.data
    first = data.FIRST_PROJECTION_YEAR_DEFAULT if years is None else years[0]
    kwargs = {
        "country_name": data.COUNTRY_NAME_DEFAULT,
        "country_initial_debt": data.COUNTRY_INITIAL_DEBT_DEFAULT,
        "first_projection_year": first,
        "growth_baseline": data.GROWTH_BASELINE_DEFAULT,
        "interest_baseline": data.INTEREST_BASELINE_DEFAULT,
        "primary_balance_baseline": data.PRIMARY_BALANCE_BASELINE_DEFAULT,
    }
    if years is None:
        return kwargs
    for name in ("growth_baseline", "interest_baseline", "primary_balance_baseline"):
        kwargs[name] = _rekey(kwargs[name], years)
    return kwargs


def test_labelled_package_exposes_runtime_axes(labelled_pkg) -> None:
    assert labelled_pkg.data.LABELLED_AXES == {"TIME_PERIOD": "engine_year_labels"}
    assert hasattr(labelled_pkg.model.Model, "cells")


def test_labeller_reaches_every_time_period_output(labelled_pkg) -> None:
    assert "first_projection_year" in required_param_names(labelled_pkg.compute_output_baseline)
    assert "first_projection_year" in required_param_names(labelled_pkg.compute_output_shocked)
    assert "engine_year_labels" in required_param_names(labelled_pkg.internals.output_baseline)
    assert "engine_year_labels" in required_param_names(labelled_pkg.internals.shock_active)


def test_snapshot_numeric_parity_and_shift_oracle(labelled_pkg) -> None:
    baseline = labelled_pkg.compute_output_baseline(**_baseline_kwargs(labelled_pkg))
    assert tuple(baseline.domain.axes[0].keys) == (1, 2, 3, 4, 5)
    assert tuple(baseline[year] for year in range(1, 6)) == pytest.approx(
        _DEFAULT_BASELINE, abs=1e-9
    )
    years = (2024, 2025, 2026, 2027, 2028)
    shifted = labelled_pkg.compute_output_baseline(**_baseline_kwargs(labelled_pkg, years=years))
    assert tuple(shifted.domain.axes[0].keys) == years
    for index, year in enumerate(years):
        assert shifted[year] == pytest.approx(_DEFAULT_BASELINE[index])
    with pytest.raises(Exception, match="unknown labels"):
        labelled_pkg.compute_output_baseline(
            **{**_baseline_kwargs(labelled_pkg), "first_projection_year": 2024}
        )


def test_model_cells_and_honest_internals(labelled_pkg) -> None:
    years = (2024, 2025, 2026, 2027, 2028)
    kwargs = _baseline_kwargs(labelled_pkg, years=years)
    kwargs.update(
        {
            "shock_year": 2025,
            "shock_type": labelled_pkg.data.SHOCK_TYPE_DEFAULT,
            "shock_magnitudes": labelled_pkg.data.SHOCK_MAGNITUDES_DEFAULT,
        }
    )
    model = labelled_pkg.model.Model(**kwargs)
    assert model.baseline_path_internal.sel(TIME_PERIOD=2026) == pytest.approx(
        model.output_baseline[2026]
    )
    cells = model.cells("shock_active")
    assert dict(cells)[(2024,)].endswith("C10")
    assert dict(cells)[(2028,)].endswith("G10")
    assert model.shock_active[2024] == 0
    assert model.shock_active[2025] == 1
