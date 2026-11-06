"""Issue 1073 — a scan formula must bind its labeller axis locally.

One recurrence group is one function. Each member used to assign `_ax_*` in
that function, and the nested formula closed over the name. Readers run
after every assignment, so every formula saw the last labeller.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_model_from_workbook import _plant_formula_cache


def _plant_years(path: Path, years: dict[str, int]) -> None:
    for coordinate, year in years.items():
        _plant_formula_cache(path, coordinate, str(year))


def _function_source(module: str, name: str) -> str:
    marker = f"def {name}"
    start = module.index(marker)
    nxt = module.find("\n\ndef ", start + 1)
    return module[start:] if nxt < 0 else module[start:nxt]


def _nested_formula(source: str, name: str) -> str:
    """Return one nested `def`, through the blank line that closes it."""
    marker = f"def {name}("
    start = source.index(marker)
    nxt = source.find("\n    def ", start + len(marker))
    blank = source.find("\n\n", start)
    end = len(source)
    for index in (nxt, blank):
        if index > start:
            end = min(end, index)
    return source[start:end]


def test_labeller_axis_lambda_stays_in_the_formula() -> None:
    """A lookup callback that reads `_ax_*` is not built in the enclosing scan."""
    from excel_grapher.exporter.inverted_tree.named_emit import _outermost_tables

    lazy = ast.parse(
        "lazy_table(((lambda: series[_ax_time_period.keys[0]],),))",
        mode="eval",
    )
    assert _outermost_tables(lazy.body, set()) == []
    eager = ast.parse("view(series, cols=_ax_time_period.keys)", mode="eval")
    hoisted = _outermost_tables(eager.body, set())
    assert len(hoisted) == 1
    assert isinstance(hoisted[0], ast.Call)


def test_scan_formulas_bind_distinct_formula_header_axes(tmp_path: Path) -> None:
    """Two formula-header windows in one scan keep their own `TIME_PERIOD` keys.

    `early` labels 2017–2019 and `late` labels 2019–2021. Neither window
    contains the other, so each input labels itself. `alpha` and `beta` read
    each other and form one scan. `alpha` at 2017 is the literal 10, which is
    position 0 on `early` and not a key of `late`.
    """
    path = tmp_path / "shared_axis.xlsx"
    write_workbook(
        path,
        {
            "Data": {
                "A1": 2017,
                "B1": "=A1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "E1": 2019,
                "F1": "=E1",
                "G1": "=F1+1",
                "H1": "=G1+1",
                "B2": 1.0,
                "C2": 2.0,
                "D2": 3.0,
                "F2": 4.0,
                "G2": 5.0,
                "H2": 6.0,
                "B4": "=10",
                "C4": "=G5",
                "D4": "=30",
                "F5": "=C4",
                "G5": "=40",
                "H5": "=B4",
            }
        },
    )
    _plant_years(
        path,
        {"B1": 2017, "C1": 2018, "D1": 2019, "F1": 2019, "G1": 2020, "H1": 2021},
    )
    document = bindings_document(
        series_entry("early", "Data!B2:D2", layout="series", direction="input", header_row=1),
        series_entry("late", "Data!F2:H2", layout="series", direction="input", header_row=1),
        series_entry("alpha", "Data!B4:D4", layout="series", direction="output", header_row=1),
        series_entry("beta", "Data!F5:H5", layout="series", direction="output", header_row=1),
        schema_version="1.17.0",
    )
    modules = generate_inverted(path, document)
    internals = modules["internals.py"]
    scan_name = next(
        line.split("(", 1)[0].removeprefix("def ")
        for line in internals.splitlines()
        if line.startswith("def scan_")
    )
    scan = _function_source(internals, scan_name)
    alpha = _nested_formula(scan, "alpha_formula")
    beta = _nested_formula(scan, "beta_formula")
    assert "_ax_time_period = early.domain.axis('TIME_PERIOD')" in alpha
    assert "_ax_time_period = late.domain.axis('TIME_PERIOD')" in beta
    assert alpha.index("_ax_time_period = ") < alpha.index(".position(")
    assert beta.index("_ax_time_period = ") < beta.index(".position(")
    assert not any(line.startswith("    _ax_") for line in scan.splitlines())

    pkg = load_package(modules, tmp_path, name="shared_axis")
    alpha_values = pkg.compute_alpha(pkg.AlphaInputs.from_defaults())
    beta_values = pkg.compute_beta(pkg.BetaInputs.from_defaults())
    assert tuple(alpha_values.domain.axes[0].keys) == (2017, 2018, 2019)
    assert tuple(beta_values.domain.axes[0].keys) == (2019, 2020, 2021)
    assert [value for _coord, value in alpha_values.items()] == pytest.approx([10.0, 40.0, 30.0])
    assert [value for _coord, value in beta_values.items()] == pytest.approx([40.0, 40.0, 10.0])


def test_hoisted_formula_header_view_binds_the_axis_before_the_view(tmp_path: Path) -> None:
    """A hoisted view names `_ax_*` in the enclosing function, before the view."""
    path = tmp_path / "sum_window.xlsx"
    write_workbook(
        path,
        {
            "Data": {
                "A1": 2017,
                "B1": "=A1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "B2": 1.0,
                "C2": 2.0,
                "D2": 3.0,
                "B4": "=SUM(B2:D2)",
                "C4": "=SUM(B2:D2)+1",
                "D4": "=C4",
            }
        },
    )
    _plant_years(path, {"B1": 2017, "C1": 2018, "D1": 2019})
    document = bindings_document(
        series_entry("early", "Data!B2:D2", layout="series", direction="input", header_row=1),
        series_entry("out", "Data!B4:D4", layout="series", direction="output", header_row=1),
        schema_version="1.17.0",
    )
    modules = generate_inverted(path, document)
    body = _function_source(modules["internals.py"], "out")
    binding = "_ax_time_period = early.domain.axis('TIME_PERIOD')"
    view = "_out_table_0 = view(early, cols=_ax_time_period.keys)"
    assert binding in body
    assert view in body
    assert body.index(binding) < body.index(view)
    formula = _nested_formula(body, "formula")
    assert binding in formula
    assert formula.index(binding) < formula.index(".position(")

    pkg = load_package(modules, tmp_path, name="sum_window")
    result = pkg.compute_out(pkg.OutInputs.from_defaults())
    assert tuple(result.domain.axes[0].keys) == (2017, 2018, 2019)
    assert [value for _coord, value in result.items()] == pytest.approx([6.0, 7.0, 7.0])


def test_deferred_scalar_binds_authored_labeller_inside_the_reader(tmp_path: Path) -> None:
    """A scalar reader in a scan binds `_ax_*` in its own function."""
    path = tmp_path / "deferred_scalar.xlsx"
    write_workbook(
        path,
        {
            "Engine": {
                "A1": "=2024",
                "B1": "=A1+1",
                "C1": "=B1+1",
                "A2": 1.0,
                "C2": 3.0,
                "D2": "=SUM(A2:C2)+A3",
                "A3": "=10",
                "B3": "=A3+1",
                "C3": "=D2",
            }
        },
    )
    _plant_years(path, {"A1": 2024, "B1": 2025, "C1": 2026})
    odd = series_entry(
        "odd_years",
        "Engine!A2:C2",
        layout="series",
        direction="input",
        header_row=1,
    )
    odd["exclude_columns"] = ["B"]
    document = bindings_document(
        series_entry(
            "year_labels",
            "Engine!A1:C1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        odd,
        series_entry("total", "Engine!D2", direction="output", dtype="float"),
        series_entry(
            "path",
            "Engine!A3:C3",
            layout="series",
            direction="output",
            header_row=1,
        ),
        schema_version="1.17.0",
    )
    modules = generate_inverted(path, document, blank_ranges=["Engine!B2"])
    internals = modules["internals.py"]
    scan = _function_source(internals, "scan_total")
    total = _nested_formula(scan, "total_formula")
    path_formula = _nested_formula(scan, "path_formula")
    binding = "_ax_time_period = year_labels.domain.axis('TIME_PERIOD')"
    assert binding in total
    assert total.index(binding) < total.index("_ax_time_period.keys")
    assert binding in path_formula
    assert path_formula.index(binding) < path_formula.index(".position(")
    assert not any(line.startswith("    _ax_") for line in scan.splitlines())

    pkg = load_package(modules, tmp_path, name="deferred_scalar_axis")
    model = pkg.Model.from_defaults()
    assert model.total == pytest.approx(14.0)
    assert [value for _coord, value in model.path.items()] == pytest.approx([10.0, 11.0, 14.0])
