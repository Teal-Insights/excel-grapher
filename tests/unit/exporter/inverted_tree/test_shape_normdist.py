"""Excel `NORMDIST` must lower through inverted-tree `excel.py` (#995)."""

from __future__ import annotations

import math
from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    invoke_public_compute,
    load_package,
    series_entry,
    write_workbook,
)

_STANDARD_DENSITY = 1.0 / math.sqrt(2.0 * math.pi)


def _normdist_workbook(tmp_path: Path, formula: str, *, z: float = 0) -> Path:
    return write_workbook(
        tmp_path / "normdist.xlsx",
        {"out": {"A1": z, "B1": formula}},
    )


def _normdist_bindings() -> dict:
    return bindings_document(
        series_entry("z", "out!A1", layout="scalar", direction="input"),
        series_entry("distress_probability", "out!B1", layout="scalar", direction="output"),
    )


def _scalar_workbook(tmp_path: Path, formula: str) -> Path:
    return write_workbook(tmp_path / "normdist_scalar.xlsx", {"out": {"A1": formula}})


def _scalar_bindings() -> dict:
    return bindings_document(
        series_entry("distress_probability", "out!A1", layout="scalar", direction="output"),
    )


def test_normdist_lowers_true_cumulative_to_xl_normdist(tmp_path: Path) -> None:
    workbook = _scalar_workbook(tmp_path, "=NORMDIST(0,0,1,TRUE)")
    modules = generate_inverted(workbook, _scalar_bindings())
    assert "xl_normdist(0, 0, 1, True)" in modules["internals.py"]
    assert "def xl_normdist" in modules["excel.py"]
    pkg = load_package(modules, tmp_path, name="normdist_true")
    assert invoke_public_compute(pkg, pkg.compute_distress_probability, {}) == pytest.approx(0.5)


def _standard_cdf(z: float) -> float:
    return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))


def test_normdist_cumulative_one_matches_evaluator(tmp_path: Path) -> None:
    formula = "=NORMDIST(A1,0,1,1)*100"
    workbook = _normdist_workbook(tmp_path, formula, z=0)
    modules = generate_inverted(workbook, _normdist_bindings())
    assert "xl_normdist(" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="normdist_cdf")
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, ["out!B1"], load_values=True)
    ).evaluate(["out!B1"])
    got = invoke_public_compute(pkg, pkg.compute_distress_probability, dict(z=0))
    assert got == pytest.approx(expected["out!B1"])
    assert got == pytest.approx(50.0)
    assert invoke_public_compute(
        pkg, pkg.compute_distress_probability, dict(z=1.96)
    ) == pytest.approx(_standard_cdf(1.96) * 100)


def test_normdist_density_and_nonpositive_stdev(tmp_path: Path) -> None:
    density = _normdist_workbook(tmp_path, "=NORMDIST(A1,0,1,FALSE)", z=0)
    pkg = load_package(
        generate_inverted(density, _normdist_bindings()),
        tmp_path,
        name="normdist_pdf",
    )
    assert invoke_public_compute(pkg, pkg.compute_distress_probability, dict(z=0)) == pytest.approx(
        _STANDARD_DENSITY
    )

    for formula, name in (
        ("=NORMDIST(A1,0,0,TRUE)", "normdist_zero_sd"),
        ("=NORMDIST(A1,0,-1,1)", "normdist_neg_sd"),
    ):
        workbook = _normdist_workbook(tmp_path, formula)
        package = load_package(
            generate_inverted(workbook, _normdist_bindings()),
            tmp_path,
            name=name,
        )
        assert (
            invoke_public_compute(package, package.compute_distress_probability, dict(z=0))
            == "#NUM!"
        )
