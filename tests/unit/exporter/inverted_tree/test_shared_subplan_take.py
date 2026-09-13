"""Issue 820 — shared helpers must own `take`, callers pass catalog-order series.

A scan computed outside a shared prefix can be windowed with `_take_after_call`
in the caller and again inside the helper. `take` fails closed, so the second
slice raises. The helper keeps the window; callers pass the unsliced series.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    required_param_names,
    series_entry,
    write_workbook,
)


def _time_series(series_id: str, data_range: str, *, direction: str) -> dict:
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction=direction,
        header_row=1,
    )


def _windowed_prefix_workbook(tmp_path: Path) -> Path:
    """5-year emp scan; shared prefix and fiscal tails live on years 2-4."""
    return write_workbook(
        tmp_path / "shared_take.xlsx",
        {
            "Inputs": {
                "B1": 2010,
                "C1": 2011,
                "D1": 2012,
                "E1": 2013,
                "F1": 2014,
                "B2": 1.0,
                "C2": 2.0,
                "D2": 3.0,
                "E2": 4.0,
                "F2": 5.0,
                "D3": 10.0,
                "E3": 20.0,
                "F3": 30.0,
                "B4": 0.0,
                "C4": 0.0,
                "D4": 0.0,
                "E4": 0.0,
                "F4": 0.0,
            },
            "Engine": {
                "B1": 2010,
                "C1": 2011,
                "D1": 2012,
                "E1": 2013,
                "F1": 2014,
                "B2": "=Inputs!B2",
                "C2": "=B2+Inputs!C2",
                "D2": "=C2+Inputs!D2",
                "E2": "=D2+Inputs!E2",
                "F2": "=E2+Inputs!F2",
                "D3": "=D2",
                "E3": "=E2",
                "F3": "=F2",
                "D4": "=D3+1",
                "E4": "=E3+1",
                "F4": "=F3+1",
                "D5": "=D4",
                "E5": "=E4",
                "F5": "=F4",
                "D6": "=D4+Inputs!D3+D2",
                "E6": "=E4+Inputs!E3+E2",
                "F6": "=F4+Inputs!F3+F2",
                "D7": "=D6+1",
                "E7": "=E6+1",
                "F7": "=F6+1",
                "B8": "=B2+Inputs!B4",
                "C8": "=C2+Inputs!C4",
                "D8": "=D2+Inputs!D4",
                "E8": "=E2+Inputs!E4",
                "F8": "=F2+Inputs!F4",
            },
        },
    )


def _windowed_prefix_bindings() -> dict:
    return bindings_document(
        _time_series("values", "Inputs!B2:F2", direction="input"),
        _time_series("extra", "Inputs!D3:F3", direction="input"),
        _time_series("other", "Inputs!B4:F4", direction="input"),
        _time_series("emp", "Engine!B2:F2", direction="internal"),
        _time_series("step0", "Engine!D3:F3", direction="internal"),
        _time_series("step1", "Engine!D4:F4", direction="internal"),
        _time_series("first", "Engine!D5:F5", direction="output"),
        _time_series("second", "Engine!D6:F6", direction="output"),
        _time_series("second_b", "Engine!D7:F7", direction="output"),
        _time_series("emp_out", "Engine!B8:F8", direction="output"),
    )


def _source(pkg, name: str, values: tuple[float, ...]):
    domain = getattr(pkg.data, name.upper() + "_DOMAIN")
    return pkg.Tensor.from_records(
        domain=domain,
        records=zip(tuple(domain), values, strict=True),
    )


def _observations(tensor, years: tuple[int, ...]) -> tuple[object, ...]:
    return tuple(tensor[year] for year in years)


def _take_before_helper_call(api: str, helper: str, series_id: str) -> bool:
    """True when `series_id = take(series_id, ...)` precedes `{helper}(` in one function."""
    tree = ast.parse(api)
    for fn in tree.body:
        if not isinstance(fn, ast.FunctionDef):
            continue
        taken = False
        for node in fn.body:
            if isinstance(node, ast.Assign) and len(node.targets) == 1:
                target = node.targets[0]
                value = node.value
                if (
                    isinstance(target, ast.Name)
                    and target.id == series_id
                    and isinstance(value, ast.Call)
                    and isinstance(value.func, ast.Name)
                    and value.func.id == "take"
                    and value.args
                    and isinstance(value.args[0], ast.Name)
                    and value.args[0].id == series_id
                ):
                    taken = True
            for call in ast.walk(node):
                if (
                    isinstance(call, ast.Call)
                    and isinstance(call.func, ast.Name)
                    and call.func.id == helper
                    and taken
                ):
                    return True
    return False


def test_shared_helper_owns_window_caller_passes_catalog_series(tmp_path: Path) -> None:
    modules = generate_inverted(_windowed_prefix_workbook(tmp_path), _windowed_prefix_bindings())
    api = modules["api.py"]
    assert "emp=self.emp" in api
    assert "take(" not in api


def test_windowed_shared_prefix_evaluates_without_double_take(tmp_path: Path) -> None:
    workbook = _windowed_prefix_workbook(tmp_path)
    document = _windowed_prefix_bindings()
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="issue820")
    values = (1.0, 2.0, 3.0, 4.0, 5.0)
    extra = (10.0, 20.0, 30.0)
    other = (0.0, 0.0, 0.0, 0.0, 0.0)
    assert required_param_names(pkg.compute_first) == ("values",)
    assert set(required_param_names(pkg.compute_second)) == {"values", "extra"}
    assert set(required_param_names(pkg.compute_emp_out)) == {"values", "other"}

    first = pkg.compute_first(values=_source(pkg, "values", values))
    second = pkg.compute_second(
        values=_source(pkg, "values", values), extra=_source(pkg, "extra", extra)
    )
    second_b = pkg.compute_second_b(
        values=_source(pkg, "values", values), extra=_source(pkg, "extra", extra)
    )
    emp_out = pkg.compute_emp_out(
        values=_source(pkg, "values", values), other=_source(pkg, "other", other)
    )
    window = (2012, 2013, 2014)
    full = (2010, 2011, 2012, 2013, 2014)
    assert _observations(first, window) == pytest.approx((7.0, 11.0, 16.0))
    assert _observations(second, window) == pytest.approx((23.0, 41.0, 61.0))
    assert _observations(second_b, window) == pytest.approx((24.0, 42.0, 62.0))
    assert _observations(emp_out, full) == pytest.approx((1.0, 3.0, 6.0, 10.0, 15.0))

    _catalog, _deps, graph = inverted_graph_parts(workbook, document)
    expected = FormulaEvaluator(graph).evaluate(
        [
            "Engine!D5",
            "Engine!E5",
            "Engine!F5",
            "Engine!D6",
            "Engine!E6",
            "Engine!F6",
            "Engine!D7",
            "Engine!E7",
            "Engine!F7",
            "Engine!B8",
            "Engine!C8",
            "Engine!D8",
            "Engine!E8",
            "Engine!F8",
        ]
    )
    assert _observations(first, window) == pytest.approx(
        tuple(expected[f"Engine!{col}5"] for col in "DEF")
    )
    assert _observations(second, window) == pytest.approx(
        tuple(expected[f"Engine!{col}6"] for col in "DEF")
    )
    assert _observations(emp_out, full) == pytest.approx(
        tuple(expected[f"Engine!{col}8"] for col in "BCDEF")
    )
