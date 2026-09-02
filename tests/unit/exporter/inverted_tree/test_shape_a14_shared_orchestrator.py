"""Layer A14 — outputs that share an engine share one orchestrator body."""

from __future__ import annotations

import inspect
import re
from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    generate_inverted,
    load_package,
    required_param_names,
)
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import (
    _a1_bindings,
    _a1_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a5_constants import (
    _a5_bindings,
    _a5_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a13_identity_flip import (
    _qcraft_bindings,
    _qcraft_workbook,
)


def test_shared_engine_emits_one_runner(tmp_path: Path) -> None:
    modules = generate_inverted(_a1_workbook(tmp_path), _a1_bindings())
    api = modules["api.py"]
    assert api.count("internals.engine_path(") == 1
    assert api.count("internals.engine_year0(") == 1
    assert "def _run_" in api
    pkg = load_package(modules, tmp_path, name="a14_a1")
    assert set(required_param_names(pkg.compute_output_path)) == {
        "initial_debt",
        "growth",
        "interest",
    }
    assert "unused_flag" not in all_param_names(pkg.compute_output_path)
    assert "unused_flag" not in all_param_names(pkg.compute_output_year1)
    path = pkg.compute_output_path(initial_debt=60.0, growth=(3.5, 3.5), interest=(4.0, 4.0))
    year1 = pkg.compute_output_year1(initial_debt=60.0, growth=(3.5, 3.5), interest=(4.0, 4.0))
    if isinstance(year1, tuple):
        year1 = year1[0]
    assert year1 == pytest.approx(path[0])


def test_disjoint_closures_keep_separate_bodies(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_a5_workbook(tmp_path), _a5_bindings()), tmp_path, name="a14_a5"
    )
    assert "shock_year" not in all_param_names(pkg.compute_output_baseline)
    assert "shock_year" in required_param_names(pkg.compute_output_shocked)
    baseline_src = inspect.getsource(pkg.compute_output_baseline)
    assert "shocked_path" not in baseline_src
    assert "shock_year" not in baseline_src
    runner = re.search(r"_run_\d+", baseline_src)
    if runner is not None:
        runner_src = inspect.getsource(getattr(pkg.api, runner.group()))
        assert "shocked_path" not in runner_src
        assert "shock_year" not in runner_src
    assert pkg.compute_output_baseline(value=10.0) == pytest.approx((10.0, 10.0))
    assert pkg.compute_output_shocked(value=10.0, shock_year=1) == pytest.approx((11.0, 11.0))


def test_identity_flip_outputs_share_one_scan_call(tmp_path: Path) -> None:
    modules = generate_inverted(_qcraft_workbook(tmp_path), _qcraft_bindings())
    api = modules["api.py"]
    assert api.count("internals.scan_") == 1
    pkg = load_package(modules, tmp_path, name="a14_qc")
    emp = pkg.compute_employment_growth()
    prod = pkg.compute_labour_productivity_growth()
    growth = pkg.compute_real_gdp_growth()
    assert emp == pytest.approx((2.9411764705882355, 3.0, 3.0))
    assert prod == pytest.approx((2.0, 0.9708737864077671, 2.0))
    assert growth == pytest.approx((5.0, 4.0, 5.06))
