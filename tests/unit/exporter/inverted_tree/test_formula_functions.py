"""Each formula series is a function of its semantic coordinates.

The generated calculation names its coordinates as parameters and returns
the formula value; the runtime applies it over the required domain and
stores Excel errors as codes. No per-cell record loop is written out.
"""

from __future__ import annotations

from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import generate_inverted, load_package
from tests.unit.exporter.inverted_tree.test_named_provenance import (
    _horizon_bindings,
    _horizon_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _zipper_bindings,
    _zipper_workbook,
)


def test_series_bodies_are_coordinate_functions(tmp_path: Path) -> None:
    modules = generate_inverted(_horizon_workbook(tmp_path, 5), _horizon_bindings(5))
    internals = modules["internals.py"]
    assert (
        "    def formula(time_period: int) -> float | str | None:\n"
        "        return as_measure(xl_mul(flow[time_period], 2))\n"
        "\n"
        "    return data.Twice.collect(evaluate(formula, data.TWICE_REQUIRED))\n"
    ) in internals
    assert "_records" not in internals
    assert "_coordinate" not in internals
    pkg = load_package(modules, tmp_path, name="formula_functions")
    assert pkg.compute_twice(flow=pkg.data.FLOW_DEFAULT)[2022] == 6.0


def test_recurrence_readers_call_coordinate_functions(tmp_path: Path) -> None:
    modules = generate_inverted(_zipper_workbook(tmp_path), _zipper_bindings())
    internals = modules["internals.py"]
    assert "    def debt_formula(time_period: int) -> float | str | None:" in internals
    assert "    debt = CoordinateReader('debt', data.DEBT_REQUIRED, debt_formula)" in internals
    assert "_coordinate" not in internals
    pkg = load_package(modules, tmp_path, name="formula_readers")
    got = pkg.compute_debt()
    assert tuple(got[year] for year in (2009, 2010, 2011)) == (100.0, 102.0, 104.04)
