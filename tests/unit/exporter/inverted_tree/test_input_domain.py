"""Issue 666 — inverted-tree `compute_*` arguments honor `input.domain`."""

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
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import (
    _a1_bindings,
    _a1_workbook,
)


def _enum_flag_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "flag.xlsx",
        {
            "Inputs": {"A1": 0},
            "Outputs": {"A1": "=Inputs!A1"},
        },
    )


def _enum_flag_bindings() -> dict:
    return bindings_document(
        series_entry(
            "flag",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="int",
            domain={"enum": [0, 1]},
        ),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
    )


def _rate_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "rate.xlsx",
        {
            "Inputs": {"B1": 0.25, "C1": 0.5, "B10": 1, "C10": 2},
            "Outputs": {"A1": "=Inputs!B1", "B1": "=Inputs!C1", "A10": 1, "B10": 2},
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
            domain={"real_between": {"min": 0, "max": 1}},
        ),
        series_entry(
            "out",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_scalar_enum_domain_accepts_and_rejects(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings()),
        tmp_path,
        name="domain_enum",
    )
    assert pkg.compute_out(flag=0) == (0,)
    assert pkg.compute_out(flag=1) == (1,)
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.compute_out(flag=2)


def test_series_real_between_names_series_on_out_of_range_member(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_rate_workbook(tmp_path), _rate_bindings()),
        tmp_path,
        name="domain_rate",
    )
    assert pkg.compute_out(rate=(0.0, 1.0)) == pytest.approx((0.0, 1.0))
    with pytest.raises(ValueError, match=r"rate\[1\] out of domain"):
        pkg.compute_out(rate=(0.0, 1.1))


def test_no_input_domain_does_not_emit_domain_guard(tmp_path: Path) -> None:
    modules = generate_inverted(_a1_workbook(tmp_path), _a1_bindings())
    assert "require_input_domain" not in modules["api.py"]
    assert "require_input_domain" not in modules["internals.py"]


def test_domain_checks_live_on_orchestrator_not_internals(tmp_path: Path) -> None:
    modules = generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings())
    assert "require_input_domain(flag" in modules["api.py"]
    assert "require_input_domain" not in modules["internals.py"]


def test_shared_runner_checks_domain_before_evaluation(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "shared.xlsx",
        {
            "Inputs": {"A1": 0},
            "Engine": {"A1": "=Inputs!A1", "B1": "=Inputs!A1+1"},
            "Outputs": {"A1": "=Engine!A1", "B1": "=Engine!B1"},
        },
    )
    document = bindings_document(
        series_entry(
            "flag",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="int",
            domain={"enum": [0, 1]},
        ),
        series_entry("engine_a", "Engine!A1", layout="scalar", direction="internal", dtype="int"),
        series_entry("engine_b", "Engine!B1", layout="scalar", direction="internal", dtype="int"),
        series_entry("out_a", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
        series_entry("out_b", "Outputs!B1", layout="scalar", direction="output", dtype="int"),
    )
    modules = generate_inverted(workbook, document)
    api = modules["api.py"]
    assert "def _run_0" in api
    assert api.count("require_input_domain(flag") >= 2
    pkg = load_package(modules, tmp_path, name="domain_shared")
    assert pkg.compute_out_a(flag=0) == (0,)
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.compute_out_b(flag=2)
