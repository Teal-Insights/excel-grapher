"""Layer A3 — `take` gathers by index; scan edges use predecessor-closure."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import _a1_bindings, _a1_workbook


def test_engine_path_requires_complete_declared_result_inputs(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a1_bindings()), tmp_path, name="a3_buf")
    year0 = pkg.internals.engine_year0(initial_debt=60.0)
    partial = pkg.Tensor.from_records(
        domain=pkg.Domain.product(pkg.Axis("TIME_PERIOD", (1,), int)), records=(((1,), 3.5),)
    )
    with pytest.raises(pkg.tensor.SchemaError, match="growth.*required coordinate.*2"):
        pkg.internals.engine_path(
            engine_year0=year0, growth=partial, interest=pkg.data.INTEREST_DEFAULT
        )


def test_misaligned_growth_interest_raise(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a1_bindings()), tmp_path, name="a3_align")
    wrong = pkg.Tensor.from_records(
        domain=pkg.Domain.product(pkg.Axis("TIME_PERIOD", (2, 3), int)),
        records=(((2,), 4.0), ((3,), 4.0)),
    )
    with pytest.raises(pkg.tensor.SchemaError, match="interest.*required coordinate.*1"):
        pkg.internals.engine_path(engine_year0=60.0, growth=pkg.data.GROWTH_DEFAULT, interest=wrong)


def test_scan_does_not_implicitly_rebase_a_later_year(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    modules = generate_inverted(workbook, _a1_bindings())
    assert "_kernels.engine_path(" not in modules["internals.py"]
    assert "CoordinateReader" in modules["internals.py"]
    assert "engine_path[time_period - 1]" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="a3_restart")
    full = pkg.internals.engine_path(
        engine_year0=60.0, growth=pkg.data.GROWTH_DEFAULT, interest=pkg.data.INTEREST_DEFAULT
    )
    later = pkg.Tensor.from_records(
        domain=pkg.Domain.product(pkg.Axis("TIME_PERIOD", (2,), int)), records=(((2,), 3.5),)
    )
    with pytest.raises(pkg.tensor.SchemaError, match="growth.*required coordinate.*1"):
        pkg.internals.engine_path(
            engine_year0=full[1], growth=later, interest=pkg.data.INTEREST_DEFAULT
        )


def test_public_compute_takes_named_series(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    modules = generate_inverted(workbook, _a1_bindings())
    assert "trim(" not in modules["api.py"]
    pkg = load_package(modules, tmp_path, name="a3_y1")
    value = pkg.compute_output_year1(
        initial_debt=60.0,
        growth=pkg.data.GROWTH_DEFAULT,
        interest=pkg.data.INTEREST_DEFAULT,
    )
    full = pkg.compute_output_path(
        initial_debt=60.0, growth=pkg.data.GROWTH_DEFAULT, interest=pkg.data.INTEREST_DEFAULT
    )
    assert value == pytest.approx(full[1])
    with pytest.raises(pkg.tensor.SchemaError, match="expected Tensor"):
        pkg.compute_output_year1(initial_debt=60.0, growth=(3.5,), interest=(4.0,))


def _middle_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a3_mid.xlsx",
        {
            "Inputs": {
                "A1": 60.0,
                "B1": 3.0,
                "C1": 3.5,
                "D1": 4.0,
                "E1": 4.5,
                "B2": 4.0,
                "C2": 4.5,
                "D2": 5.0,
                "E2": 5.5,
                "B10": 1,
                "C10": 2,
                "D10": 3,
                "E10": 4,
            },
            "Engine": {
                "B1": "=Inputs!A1",
                "C1": "=B1*(1+Inputs!B2/100)/(1+Inputs!B1/100)",
                "D1": "=C1*(1+Inputs!C2/100)/(1+Inputs!C1/100)",
                "E1": "=D1*(1+Inputs!D2/100)/(1+Inputs!D1/100)",
                "F1": "=E1*(1+Inputs!E2/100)/(1+Inputs!E1/100)",
                "C10": 1,
                "D10": 2,
                "E10": 3,
                "F10": 4,
            },
            "Outputs": {
                "A1": "=Engine!D1",
                "B1": "=Engine!E1",
                "A10": 2,
                "B10": 3,
            },
        },
    )


def _middle_bindings() -> dict:
    return bindings_document(
        series_entry("initial_debt", "Inputs!A1", layout="scalar", direction="input"),
        series_entry(
            "growth",
            "Inputs!B1:E1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "interest",
            "Inputs!B2:E2",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry("engine_year0", "Engine!B1", layout="scalar", direction="internal"),
        series_entry(
            "engine_path",
            "Engine!C1:F1",
            layout="series",
            direction="internal",
            header_row=10,
        ),
        series_entry(
            "output_mid",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_middle_slice_scan_uses_predecessor_closure(tmp_path: Path) -> None:
    workbook = _middle_workbook(tmp_path)
    modules = generate_inverted(workbook, _middle_bindings())
    api = modules["api.py"]
    assert "trim(" not in api
    assert "growth: data.Growth[" in api
    assert "interest: data.Interest[" in api
    pkg = load_package(modules, tmp_path, name="a3_mid")
    growth = pkg.data.GROWTH_DEFAULT
    interest = pkg.data.INTEREST_DEFAULT
    year0 = pkg.internals.engine_year0(initial_debt=60.0)
    full = pkg.internals.engine_path(engine_year0=year0, growth=growth, interest=interest)
    got = pkg.compute_output_mid(initial_debt=60.0, growth=growth, interest=interest)
    assert (got[2], got[3]) == pytest.approx((full[2], full[3]))
    graph = create_dependency_graph(workbook, ["Outputs!A1", "Outputs!B1"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1", "Outputs!B1"])
    assert (got[2], got[3]) == pytest.approx((expected["Outputs!A1"], expected["Outputs!B1"]))


def _punched_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a3_punch.xlsx",
        {
            "Inputs": {
                "B1": 10.0,
                "C1": 20.0,
                "D1": 30.0,
                "B10": 1,
                "C10": 2,
                "D10": 3,
            },
            "Engine": {
                "B1": "=Inputs!B1+1",
                "C1": "=Inputs!C1+1",
                "D1": "=Inputs!D1+1",
                "B10": 1,
                "C10": 2,
                "D10": 3,
            },
            "Outputs": {
                "A1": "=Engine!B1",
                "B1": "=Engine!D1",
                "A10": 1,
                "B10": 3,
            },
        },
    )


def _punched_bindings() -> dict:
    return bindings_document(
        series_entry(
            "values",
            "Inputs!B1:D1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "engine_plus",
            "Engine!B1:D1",
            layout="series",
            direction="internal",
            header_row=10,
        ),
        series_entry(
            "output_punched",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_punched_elementwise_gathers_holes(tmp_path: Path) -> None:
    workbook = _punched_workbook(tmp_path)
    modules = generate_inverted(workbook, _punched_bindings())
    api = modules["api.py"]
    assert "trim(" not in api
    assert "values: data.Values[" in api
    pkg = load_package(modules, tmp_path, name="a3_punch")
    got = pkg.compute_output_punched(values=pkg.data.VALUES_DEFAULT)
    assert tuple(got.domain) == ((1,), (3,))
    assert (got[1], got[3]) == pytest.approx((11.0, 31.0))
    with pytest.raises(pkg.tensor.SchemaError, match="expected Tensor"):
        pkg.compute_output_punched(values=(10.0, 30.0))
