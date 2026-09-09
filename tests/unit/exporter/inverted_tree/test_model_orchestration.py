"""Public outputs evaluate through one memoized `Model` of named formulas.

Every formula series is defined exactly once in `api.py` as a lazily
evaluated attribute of `Model`. A public `compute_*` function binds only its
own leaves and reads one attribute, so shared prefixes are never re-emitted
per output and no output evaluates another output's private tail.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    generate_inverted,
    load_package,
    required_param_names,
)
from tests.unit.exporter.inverted_tree.test_shape_a13_identity_flip import (
    _qcraft_bindings,
    _qcraft_workbook,
)
from tests.unit.exporter.inverted_tree.test_shared_subplan import (
    _PREFIX_LEN,
    _observations,
    _prefix_bindings,
    _prefix_workbook,
    _source,
)


def test_every_formula_is_one_model_attribute(tmp_path: Path) -> None:
    modules = generate_inverted(_prefix_workbook(tmp_path), _prefix_bindings())
    api = modules["api.py"]
    assert "class Model:" in api
    assert "def _shared_" not in api
    assert "def _run_" not in api
    for index in range(_PREFIX_LEN):
        assert api.count(f"internals.step_{index}(") == 1
        assert api.count(f"    def step_{index}(self)") == 1
    first_fn = api[api.index("def compute_first") :].split("\ndef ")[0]
    assert "extra" not in first_fn
    assert "internals." not in first_fn
    assert ".first" in first_fn


def test_model_evaluates_lazily_and_per_instance(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_prefix_workbook(tmp_path), _prefix_bindings()),
        tmp_path,
        name="model_lazy",
    )
    assert required_param_names(pkg.compute_first) == ("values",)
    values = (10.0, 20.0, 30.0)
    model = pkg.api.Model(values=_source(pkg, "values", values))
    assert "step_0" not in vars(model)
    first = model.first
    assert "step_0" in vars(model)
    assert model.first is first
    assert _observations(first) == pytest.approx(tuple(v + _PREFIX_LEN + 1 for v in values))
    other = pkg.api.Model(values=_source(pkg, "values", (11.0, 20.0, 30.0)))
    assert other.first[2020] == pytest.approx(first[2020] + 1.0)
    assert _observations(pkg.compute_first(values=_source(pkg, "values", values))) == (
        pytest.approx(_observations(first))
    )


def test_recurrence_groups_use_short_scan_names(tmp_path: Path) -> None:
    modules = generate_inverted(_qcraft_workbook(tmp_path), _qcraft_bindings())
    api = modules["api.py"]
    assert api.count("internals.scan_") == 1
    names = re.findall(r"internals\.(scan_\w+)\(", api)
    assert names and all(len(name) < 80 for name in names)
    internals = modules["internals.py"]
    assert f"def {names[0]}(" in internals
