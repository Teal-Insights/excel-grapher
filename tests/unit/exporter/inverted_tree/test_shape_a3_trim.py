"""Layer A3 — trim does not expand the leaf set."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    generate_inverted,
    load_package,
)
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import _a1_bindings, _a1_workbook


def test_engine_path_accepts_year1_prefix(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a1_bindings()), tmp_path, name="a3_trim")
    year0 = pkg.internals.engine_year0(60.0)
    one = pkg.internals.engine_path(year0, (3.5,), (4.0,))
    assert len(one) == 1
    full = pkg.internals.engine_path(year0, (3.5, 3.5), (4.0, 4.0))
    assert one[0] == pytest.approx(full[0])


def test_misaligned_growth_interest_raise(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a1_bindings()), tmp_path, name="a3_align")
    year0 = pkg.internals.engine_year0(60.0)
    with pytest.raises(ValueError, match="misaligned"):
        pkg.internals.engine_path(year0, (3.5, 3.5), (4.0,))


def test_prefix_trim_restart_from_year1_debt(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a1_bindings()), tmp_path, name="a3_restart")
    year0 = pkg.internals.engine_year0(60.0)
    full = pkg.internals.engine_path(year0, (3.5, 3.5), (4.0, 4.0))
    restarted = pkg.internals.engine_path(full[0], (3.5,), (4.0,))
    assert restarted[0] == pytest.approx(full[1])


def test_compute_year1_does_not_require_year2(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a1_bindings()), tmp_path, name="a3_y1")
    value = pkg.compute_output_year1(initial_debt=60.0, growth=(3.5,), interest=(4.0,))
    if isinstance(value, tuple):
        assert len(value) == 1
        value = value[0]
    full = pkg.compute_output_path(initial_debt=60.0, growth=(3.5, 3.5), interest=(4.0, 4.0))
    assert value == pytest.approx(full[0])
