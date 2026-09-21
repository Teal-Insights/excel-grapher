"""Public outputs evaluate through one memoized `Model` of named formulas.

Every formula series is defined exactly once in `model.py` as a lazily
evaluated attribute of `Model`. A public `compute_*` function binds only its
own leaves and reads one attribute, so shared prefixes are never re-emitted
per output and no output evaluates another output's private tail.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    input_field_names,
    invoke_public_compute,
    load_package,
    series_entry,
    write_workbook,
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


def test_generated_package_exports_model(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "model_export.xlsx",
        {"Inputs": {"A1": 2.0, "B1": "=A1*3"}},
    )
    document = bindings_document(
        series_entry("seed", "Inputs!A1"),
        series_entry("result", "Inputs!B1", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="model_export")
    assert "Model" in pkg.__all__
    assert pkg.Model is pkg.model.Model
    assert not hasattr(pkg.api, "Model")


def test_every_formula_is_one_model_attribute(tmp_path: Path) -> None:
    modules = generate_inverted(_prefix_workbook(tmp_path), _prefix_bindings())
    model = modules["model.py"]
    api = modules["api.py"]
    assert "class Model(_BoundInputs):" in model
    assert "class Model:" not in api
    assert "def _shared_" not in model
    assert "def _run_" not in model
    for index in range(_PREFIX_LEN):
        assert model.count(f"internals.step_{index}(") == 1
        assert model.count(f"    def step_{index}(self)") == 1
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
    assert input_field_names(pkg, pkg.compute_first) == ("values",)
    values = (10.0, 20.0, 30.0)
    model = pkg.model.Model(values=_source(pkg, "values", values))
    assert "step_0" not in vars(model)
    first = model.first
    assert "step_0" in vars(model)
    assert model.first is first
    assert _observations(first) == pytest.approx(tuple(v + _PREFIX_LEN + 1 for v in values))
    other = pkg.model.Model(values=_source(pkg, "values", (11.0, 20.0, 30.0)))
    assert other.first[2020] == pytest.approx(first[2020] + 1.0)
    assert _observations(
        invoke_public_compute(pkg, pkg.compute_first, dict(values=_source(pkg, "values", values)))
    ) == (pytest.approx(_observations(first)))


def test_recurrence_groups_use_short_scan_names(tmp_path: Path) -> None:
    modules = generate_inverted(_qcraft_workbook(tmp_path), _qcraft_bindings())
    model = modules["model.py"]
    assert model.count("internals.scan_") == 1
    names = re.findall(r"internals\.(scan_\w+)\(", model)
    assert names and all(len(name) < 80 for name in names)
    internals = modules["internals.py"]
    assert f"def {names[0]}(" in internals
