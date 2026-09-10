"""Generated series classes carry their value type and schema, so call sites stay short."""

from __future__ import annotations

from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import generate_inverted, load_package
from tests.unit.exporter.inverted_tree.test_named_provenance import (
    _horizon_bindings,
    _horizon_workbook,
)


def test_facades_are_concrete_and_collect_their_own_records(tmp_path: Path) -> None:
    modules = generate_inverted(_horizon_workbook(tmp_path, 5), _horizon_bindings(5))
    data = modules["data.py"]
    internals = modules["internals.py"]
    assert "class Twice(Series[float | str | None]):\n" in data
    assert "FLOW_DEFAULT = Flow(FLOW_DOMAIN, " in data
    assert "def twice(*, flow: data.Flow) -> data.Twice:" in internals
    assert "@publish(data.TWICE_SCHEMA, cells=data.TWICE_CELLS)" in internals
    assert "    return data.Twice.collect(evaluate(formula, data.TWICE_REQUIRED))" in internals
    assert "[float | str | None]" not in internals
    pkg = load_package(modules, tmp_path, name="concrete_facades")
    twice = pkg.compute_twice(flow=pkg.data.FLOW_DEFAULT)
    assert isinstance(twice, pkg.data.Twice)
    assert pkg.compute_twice.__key__ == ("TIME_PERIOD",)
    assert pkg.compute_twice.__domain__ is pkg.data.TWICE_REQUIRED
    assert pkg.internals.twice.__key__ == ("TIME_PERIOD",)
    assert twice[2022] == 6.0


def test_schemas_share_one_value_type_tuple_per_dtype(tmp_path: Path) -> None:
    modules = generate_inverted(_horizon_workbook(tmp_path, 5), _horizon_bindings(5))
    data = modules["data.py"]
    assert "FLOAT_VALUES = (int, float, bool, str, type(None))\n" in data
    assert "TWICE_SCHEMA = TensorSchema('twice', TWICE_REQUIRED, FLOAT_VALUES)\n" in data
    assert data.count("type(None)") == 1
