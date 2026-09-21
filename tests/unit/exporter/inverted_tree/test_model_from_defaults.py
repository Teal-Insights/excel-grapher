"""Issue 932 — `Model.from_defaults` binds `data.*_DEFAULT` in one call."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_input_domain import (
    _enum_flag_bindings,
    _enum_flag_workbook,
)


def _seed_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "from_defaults.xlsx",
        {"Inputs": {"A1": 2.0, "B1": "=A1*3"}},
    )


def _seed_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("result", "Inputs!B1", layout="scalar", direction="output"),
    )


def _constant_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "from_defaults_const.xlsx",
        {
            "Inputs": {"A1": 2.0, "B1": 3.0},
            "Outputs": {"A1": "=Inputs!A1*Inputs!B1"},
        },
    )


def _constant_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("factor", "Inputs!B1", layout="scalar", direction="constant"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )


def test_from_defaults_emits_explicit_input_kwargs(tmp_path: Path) -> None:
    model = generate_inverted(_seed_workbook(tmp_path), _seed_bindings())["model.py"]
    assert "cls.__annotations__" not in model
    assert "seed: float | str = data.SEED_DEFAULT" in model
    assert "return cls(seed=seed)" in model
    assert "unknown = inputs.keys() -" in model


def test_from_defaults_binds_snapshot_and_overrides(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_seed_workbook(tmp_path), _seed_bindings()),
        tmp_path,
        name="from_defaults_seed",
    )
    Model = pkg.Model
    assert pkg.data.SEED_DEFAULT == 2.0
    empty = Model()
    with pytest.raises(AttributeError):
        _ = empty.result
    assert Model.from_defaults().result == 6.0
    assert Model.from_defaults(seed=4.0).result == 12.0


def test_model_constructor_rejects_unknown_inputs(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_seed_workbook(tmp_path), _seed_bindings()),
        tmp_path,
        name="from_defaults_unknown",
    )
    Model = pkg.Model
    with pytest.raises(TypeError, match="unknown inputs"):
        Model(discount_rate=0.04)
    with pytest.raises(TypeError, match="unexpected keyword argument"):
        Model.from_defaults(discount_rate=0.04)


def test_from_defaults_leaves_constants_on_data(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_constant_workbook(tmp_path), _constant_bindings()),
        tmp_path,
        name="from_defaults_const",
    )
    Model = pkg.Model
    assert pkg.data.FACTOR == 3.0
    assert not hasattr(pkg.data, "FACTOR_DEFAULT")
    assert Model.from_defaults().result == 6.0
    assert Model.from_defaults(seed=4.0).result == 12.0
    with pytest.raises(TypeError, match="unknown inputs"):
        Model(factor=5.0)
    with pytest.raises(TypeError, match="unexpected keyword argument"):
        Model.from_defaults(factor=5.0)
    with pkg.data.overrides(FACTOR=5.0):
        assert Model.from_defaults(seed=4.0).result == 20.0


def test_from_defaults_uses_init_validation_checks(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings()),
        tmp_path,
        name="from_defaults_flag",
    )
    Model = pkg.Model
    assert Model.from_defaults().out == 0
    assert Model.from_defaults(flag=1).out == 1
    with pytest.raises(ValueError, match="flag out of domain"):
        Model.from_defaults(flag=2)
    with pytest.raises(ValueError, match="flag out of domain"):
        Model(flag=2)
