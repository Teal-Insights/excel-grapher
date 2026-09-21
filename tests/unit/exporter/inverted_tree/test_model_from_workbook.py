"""Issue 942 — `Model.from_workbook` binds input leaves from a vintage xlsx."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    named_input_kwargs,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_dynamic_axis_labels import (
    _labelled_bindings,
    _labelled_workbook,
)
from tests.unit.exporter.inverted_tree.test_input_domain import (
    _enum_flag_bindings,
    _enum_flag_workbook,
)


def _seed_workbook(path: Path, seed: float) -> Path:
    return write_workbook(path, {"Inputs": {"A1": seed, "B1": "=A1*3"}})


def _seed_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("result", "Inputs!B1", layout="scalar", direction="output"),
    )


def _constant_workbook(path: Path, *, seed: float, factor: float) -> Path:
    return write_workbook(
        path,
        {
            "Inputs": {"A1": seed, "B1": factor},
            "Outputs": {"A1": "=Inputs!A1*Inputs!B1"},
        },
    )


def _constant_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("factor", "Inputs!B1", layout="scalar", direction="constant"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )


def _rate_workbook(path: Path, left: float, right: float) -> Path:
    return write_workbook(
        path,
        {
            "Inputs": {"B1": left, "C1": right, "B10": 1, "C10": 2},
            "Outputs": {"A1": "=Inputs!B1+Inputs!C1", "B1": 0, "A10": 1, "B10": 2},
        },
    )


def _rate_bindings() -> dict:
    return bindings_document(
        series_entry(
            "rate",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "out",
            "Outputs!A1",
            layout="scalar",
            direction="output",
        ),
    )


def test_from_workbook_emits_mixin_and_catalog_input_ids(tmp_path: Path) -> None:
    modules = generate_inverted(_seed_workbook(tmp_path / "template.xlsx", 2.0), _seed_bindings())
    api = modules["api.py"]
    assert "class _BoundInputs:" in api
    assert "def from_workbook" in api
    assert "cls.__annotations__" not in api
    assert '_INPUT_IDS: tuple[str, ...] = ("seed",)' in api
    assert "class Model(_BoundInputs):" in api
    assert "workbook.py" in modules
    assert "def compute_result" in api
    assert "workbook" not in api[api.index("def compute_result") :]


def test_from_workbook_reads_case_workbook_and_overrides_win(tmp_path: Path) -> None:
    template = _seed_workbook(tmp_path / "template.xlsx", 2.0)
    pkg = load_package(
        generate_inverted(template, _seed_bindings()),
        tmp_path,
        name="from_workbook_seed",
    )
    Model = pkg.api.Model
    case = _seed_workbook(tmp_path / "case.xlsx", 4.0)
    extra = write_workbook(
        tmp_path / "extra_sheets.xlsx",
        {"Inputs": {"A1": 4.0, "B1": "=A1*3"}, "Notes": {"Z9": "ignored"}},
    )

    assert pkg.data.SEED_DEFAULT == 2.0
    assert Model.from_workbook(case).result == 12.0
    assert Model.from_workbook(extra).result == 12.0
    assert Model.from_workbook(case, seed=5.0).result == 15.0
    assert pkg.compute_result(seed=Model.from_workbook(case).seed) == 12.0


def test_from_workbook_missing_input_sheet_raises(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_seed_workbook(tmp_path / "template.xlsx", 2.0), _seed_bindings()),
        tmp_path,
        name="from_workbook_missing",
    )
    missing = write_workbook(tmp_path / "missing.xlsx", {"Other": {"A1": 4.0}})
    with pytest.raises(KeyError, match="Inputs!A1"):
        pkg.api.Model.from_workbook(missing)


def test_from_workbook_rejects_unknown_overrides(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_seed_workbook(tmp_path / "template.xlsx", 2.0), _seed_bindings()),
        tmp_path,
        name="from_workbook_unknown",
    )
    case = _seed_workbook(tmp_path / "case.xlsx", 4.0)
    with pytest.raises(TypeError, match="unknown inputs"):
        pkg.api.Model.from_workbook(case, discount_rate=0.04)


def test_from_workbook_leaves_constants_on_codegen_snapshot(tmp_path: Path) -> None:
    template = _constant_workbook(tmp_path / "template.xlsx", seed=2.0, factor=3.0)
    pkg = load_package(
        generate_inverted(template, _constant_bindings()),
        tmp_path,
        name="from_workbook_const",
    )
    Model = pkg.api.Model
    case = _constant_workbook(tmp_path / "case.xlsx", seed=4.0, factor=9.0)
    assert pkg.data.FACTOR == 3.0
    assert Model.from_workbook(case).result == 12.0
    with pytest.raises(TypeError, match="unknown inputs"):
        Model.from_workbook(case, factor=5.0)
    with pkg.data.overrides(FACTOR=5.0):
        assert Model.from_workbook(case).result == 20.0


def test_from_workbook_builds_tensors_with_records(tmp_path: Path) -> None:
    template = _rate_workbook(tmp_path / "template.xlsx", 0.25, 0.5)
    pkg = load_package(
        generate_inverted(template, _rate_bindings()),
        tmp_path,
        name="from_workbook_rate",
    )
    case = _rate_workbook(tmp_path / "case.xlsx", 1.5, 2.5)
    model = pkg.api.Model.from_workbook(case)
    assert model.out == pytest.approx(4.0)
    overlay = pkg.data.RATE.with_records((((1,), 3.0), ((2,), 4.0)))
    assert pkg.api.Model.from_workbook(case, rate=overlay).out == pytest.approx(7.0)


def test_from_workbook_uses_init_validation_checks(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings()),
        tmp_path,
        name="from_workbook_flag",
    )
    Model = pkg.api.Model
    valid = write_workbook(
        tmp_path / "flag_case.xlsx",
        {"Inputs": {"A1": 1}, "Outputs": {"A1": "=Inputs!A1"}},
    )
    invalid = write_workbook(
        tmp_path / "flag_bad.xlsx",
        {"Inputs": {"A1": 2}, "Outputs": {"A1": "=Inputs!A1"}},
    )
    assert Model.from_workbook(valid).out == 1
    with pytest.raises(ValueError, match="flag out of domain"):
        Model.from_workbook(invalid)


def test_from_workbook_reuses_deferred_runtime_labeller_init(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, _labelled_bindings()),
        tmp_path,
        name="from_workbook_labelled",
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, _labelled_bindings())
    expected = call_compute(pkg, "path", named_input_kwargs(pkg, catalog, graph))
    model = pkg.api.Model.from_workbook(workbook)
    assert list(model.path.items()) == list(expected.items())
    shifted = _labelled_workbook(tmp_path, first=2025, shock_year=2026)
    with pytest.raises(pkg.tensor.SchemaError):
        pkg.api.Model.from_workbook(shifted)
