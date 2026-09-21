"""#890 None short-circuit must not change AVERAGE vs FormulaEvaluator.

When first/middle/last samples of a copy family match, unsampled members whose
affine image is missing from the producer used to record `as_measure(None)`
instead of lowering the member. AVERAGE skips `None` and includes Excel's 0
from `=blank` / `=blank+0` / `=+blank`, so export and the evaluator diverged
at named-differential tolerances (issue 892). Off-axis numeric formula copies
now lower to `0`.
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

from fastpyxl import load_workbook
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    inverted_graph_parts,
    invoke_public_compute,
    load_package,
    named_input_kwargs,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_off_axis_producer_keys import (
    _ALPHA,
    _BUCKET_ORIGIN,
    _N_HOST,
    _ORIGIN,
    _letter,
    _matrix_copy_bindings,
    _matrix_copy_workbook,
    _mcve_blank_ranges,
    _series_copy_bindings,
)

# scripts/named_differential.py
_ATOL = 1e-6
_RTOL = 1e-12
_PRODUCER_LIFE = 11


def _named_close(actual: object, expected: object, *, label: str) -> None:
    """Assert numeric agreement at named-differential tolerances."""
    if isinstance(actual, bool) or isinstance(expected, bool):
        raise AssertionError(f"{label}: unexpected bool {actual!r} vs {expected!r}")
    if not isinstance(actual, (int, float)) or not isinstance(expected, (int, float)):
        raise AssertionError(f"{label}: {actual!r} vs {expected!r}")
    left, right = float(actual), float(expected)
    assert math.isclose(left, right, rel_tol=_RTOL, abs_tol=_ATOL), (
        f"{label}: {left!r} vs {right!r} abs={abs(left - right)}"
    )


def _evaluate_cell(workbook: Path, cell: str, *, blank_ranges: tuple[str, ...]) -> object:
    graph = create_dependency_graph(workbook, [cell], load_values=True, blank_ranges=blank_ranges)
    return FormulaEvaluator(graph, blank_ranges=blank_ranges).evaluate(cell)


def _copy_average_workbook(tmp_path: Path, *, host_formula: str, name: str) -> Path:
    """1-D host copies a shorter STEP producer; A5 averages the host row."""
    cells: dict[str, object] = {}
    for life in range(_N_HOST):
        letter = _letter(life)
        cells[f"{letter}1"] = life
        cells[f"{letter}3"] = _BUCKET_ORIGIN + life
        cells[f"{letter}4"] = host_formula.format(src=letter)
        if life < _PRODUCER_LIFE:
            cells[f"{letter}2"] = float(life + 1)
    cells["A5"] = f"=AVERAGE(B4:{_letter(_N_HOST - 1)}4)"
    return write_workbook(tmp_path / name, {"Sheet": cells})


def _copy_average_bindings() -> dict[str, Any]:
    document = _series_copy_bindings()
    document["series"].append(series_entry("avg", "Sheet!A5", layout="scalar", direction="output"))
    return document


def _copy_blank_ranges() -> tuple[str, ...]:
    return (f"Sheet!{_letter(4)}2:{_letter(_N_HOST - 1)}2",)


def _exported_avg(
    workbook: Path,
    document: dict[str, Any],
    tmp_path: Path,
    name: str,
    *,
    blank_ranges: tuple[str, ...],
) -> object:
    catalog, _deps, graph = inverted_graph_parts(workbook, document, blank_ranges=blank_ranges)
    modules = generate_inverted(workbook, document, blank_ranges=blank_ranges)
    pkg = load_package(modules, tmp_path, name=name)
    return call_compute(pkg, "avg", named_input_kwargs(pkg, catalog, graph))


def _interior_hole_workbook(tmp_path: Path, *, wrap: str) -> Path:
    """Host copies a 1-D producer whose interior year is a structural blank."""
    years = tuple(range(2020, 2029))
    hole = 2022
    cells: dict[str, object] = {}
    for offset, year in enumerate(years):
        letter = get_column_letter(offset + 2)
        cells[f"{letter}1"] = year
        cells[f"{letter}3"] = f"=Sheet!{letter}2{wrap}"
        if year != hole:
            cells[f"{letter}2"] = 0.05 * (offset + 1)
    last = get_column_letter(len(years) + 1)
    cells["A4"] = f"=AVERAGE(B3:{last}3)"
    return write_workbook(tmp_path / f"interior_hole{wrap or '_bare'}.xlsx", {"Sheet": cells})


def _interior_hole_bindings(years: int = 9) -> dict[str, Any]:
    last = get_column_letter(years + 1)
    return bindings_document(
        series_entry(
            "steps", f"Sheet!B2:{last}2", layout="series", direction="input", header_row=1
        ),
        series_entry(
            "copied", f"Sheet!B3:{last}3", layout="series", direction="internal", header_row=1
        ),
        series_entry("avg", "Sheet!A4", layout="scalar", direction="output"),
        schema_version="1.16.0",
    )


def test_average_of_wrapped_off_axis_copies_matches_evaluator(tmp_path: Path) -> None:
    """`=producer+0` off-axis years stay in AVERAGE (xl_add(None, 0) is 0)."""
    workbook = _copy_average_workbook(
        tmp_path, host_formula="=Sheet!{src}2+0", name="wrapped_avg.xlsx"
    )
    blanks = _copy_blank_ranges()
    got = _exported_avg(
        workbook, _copy_average_bindings(), tmp_path, "wrapped_avg", blank_ranges=blanks
    )
    want = _evaluate_cell(workbook, "Sheet!A5", blank_ranges=blanks)
    _named_close(got, want, label="wrapped AVERAGE")


def test_average_of_unary_plus_off_axis_copies_matches_evaluator(tmp_path: Path) -> None:
    """LIC-DSF `=+producer` off-axis members lower to 0, not AVERAGE-skipped None."""
    workbook = _copy_average_workbook(
        tmp_path, host_formula="=+Sheet!{src}2", name="unary_avg.xlsx"
    )
    blanks = _copy_blank_ranges()
    got = _exported_avg(
        workbook, _copy_average_bindings(), tmp_path, "unary_avg", blank_ranges=blanks
    )
    want = _evaluate_cell(workbook, "Sheet!A5", blank_ranges=blanks)
    _named_close(got, want, label="unary-plus AVERAGE")


def test_interior_hole_wrapped_copy_average_matches_evaluator(tmp_path: Path) -> None:
    """On-axis hole copied as `=+0` is already 0 via `xl_add`; AVERAGE includes it."""
    workbook = _interior_hole_workbook(tmp_path, wrap="+0")
    blanks = ("Sheet!D2",)
    got = _exported_avg(
        workbook,
        _interior_hole_bindings(),
        tmp_path,
        "interior_hole_avg",
        blank_ranges=blanks,
    )
    want = _evaluate_cell(workbook, "Sheet!A4", blank_ranges=blanks)
    _named_close(got, want, label="interior-hole AVERAGE")


def test_average_of_unwrapped_interior_hole_copies_matches_evaluator(tmp_path: Path) -> None:
    """On-axis producer hole copied as `=blank` must be 0 in AVERAGE, not skipped None."""
    workbook = _interior_hole_workbook(tmp_path, wrap="")
    blanks = ("Sheet!D2",)
    got = _exported_avg(
        workbook,
        _interior_hole_bindings(),
        tmp_path,
        "interior_hole_unwrapped_avg",
        blank_ranges=blanks,
    )
    want = _evaluate_cell(workbook, "Sheet!A4", blank_ranges=blanks)
    _named_close(got, want, label="unwrapped interior-hole AVERAGE")


def test_average_of_on_axis_hole_formula_copies_matches_evaluator(tmp_path: Path) -> None:
    """AVERAGE of Beta includes on-axis head holes the host copies with `=Input!X`."""
    workbook = _matrix_copy_workbook(tmp_path)
    last = _letter(_N_HOST - 1)
    book = load_workbook(workbook)
    book["Host"]["A4"] = f"=AVERAGE(B3:{last}3)"
    book.save(workbook)
    book.close()
    document = _matrix_copy_bindings()
    document["series"].append(series_entry("avg", "Host!A4", layout="scalar", direction="output"))
    blanks = _mcve_blank_ranges()
    got = _exported_avg(workbook, document, tmp_path, "beta_hole_avg", blank_ranges=blanks)
    want = _evaluate_cell(workbook, "Host!A4", blank_ranges=blanks)
    _named_close(got, want, label="Beta on-axis-hole AVERAGE")


def test_isblank_of_producer_hole_stays_true(tmp_path: Path) -> None:
    """Measure-boundary `=blank` → 0 must not make `ISBLANK(producer_hole)` false."""
    workbook = _matrix_copy_workbook(tmp_path)
    book = load_workbook(workbook)
    book["Host"]["Z1"] = "=ISBLANK(Input!B3)"
    book.save(workbook)
    book.close()
    document = _matrix_copy_bindings()
    document["series"].append(
        series_entry("probe", "Host!Z1", layout="scalar", direction="output", dtype="bool")
    )
    blanks = _mcve_blank_ranges()
    catalog, _deps, graph = inverted_graph_parts(workbook, document, blank_ranges=blanks)
    modules = generate_inverted(workbook, document, blank_ranges=blanks)
    pkg = load_package(modules, tmp_path, name="isblank_hole")
    got = call_compute(pkg, "probe", named_input_kwargs(pkg, catalog, graph))
    want = _evaluate_cell(workbook, "Host!Z1", blank_ranges=blanks)
    assert got is True
    assert want is True


def test_wrapped_off_axis_member_matches_evaluator_zero(tmp_path: Path) -> None:
    """An unsampled `=Input+0` year whose life image is missing evaluates to 0."""
    n_host, matrix_life, alpha_blank_from = 5, 4, 3
    workbook = _matrix_copy_workbook(
        tmp_path,
        n_host=n_host,
        matrix_life=matrix_life,
        wrap_add_zero=True,
        name="wrapped_zero.xlsx",
    )
    document = _matrix_copy_bindings(n_host=n_host, matrix_life=matrix_life)
    blanks = (
        f"Input!{_letter(alpha_blank_from)}2:{_letter(n_host - 1)}2",
        f"Input!{_letter(0)}3:{_letter(1)}3",
        f"Input!{_letter(matrix_life)}3:{_letter(n_host - 1)}3",
    )
    modules = generate_inverted(workbook, document, blank_ranges=blanks)
    pkg = load_package(modules, tmp_path, name="wrapped_zero")
    got = invoke_public_compute(
        pkg, pkg.compute_schedule, dict(principal=pkg.data.PRINCIPAL_DEFAULT)
    )
    tail_year = _ORIGIN + n_host - 1
    tail_cell = f"Host!{_letter(n_host - 1)}2"
    want = _evaluate_cell(workbook, tail_cell, blank_ranges=blanks)
    _named_close(got[_ALPHA, tail_year], want, label=f"Alpha {tail_year}")


def test_as_measure_call_coerces_numeric_whole_body_none_to_zero() -> None:
    """`None`→`0` lives on `python_dtype` (`number` already maps to `float`)."""
    from types import SimpleNamespace

    from excel_grapher.exporter.inverted_tree.ast_emit import _as_measure_call
    from excel_grapher.exporter.inverted_tree.catalog import _DTYPE_READ

    assert _DTYPE_READ["number"] == "float"
    assert _as_measure_call("None", SimpleNamespace(python_dtype="float")) == "as_measure(0)"
    assert _as_measure_call("None", SimpleNamespace(python_dtype="int")) == "as_measure(0, 'int')"
    assert (
        _as_measure_call("None", SimpleNamespace(python_dtype="bool")) == "as_measure(None, 'bool')"
    )
    nested = "xl_add(None, 0)"
    assert (
        _as_measure_call(nested, SimpleNamespace(python_dtype="float")) == f"as_measure({nested})"
    )


def test_yoy_average_folds_deflator_ratio_with_unsampled_off_axis_year(tmp_path: Path) -> None:
    """A1-like AVERAGE must fold `deflator[t]/deflator[t-1]` and lower 2017.

    Growth spans 2015–2022 so first/middle/last (2015/2019/2022) share the lagged
    ratio. 2017 is excluded from the deflator axis (`exclude_columns` plus
    `blank_ranges`), so copying the sampled lookup would emit `as_measure(None)`
    and AVERAGE would skip that year. Lowering yields
    `IF(ISNUMBER(2016), (None/2016-1)*100, ...)` = -100, which AVERAGE includes.
    """
    years = tuple(range(2014, 2023))
    cells: dict[str, object] = {}
    hole_year = 2017
    for offset, year in enumerate(years):
        letter = get_column_letter(offset + 2)
        cells[f"{letter}1"] = year
        if year != hole_year:
            cells[f"{letter}2"] = 100.0 * (1.03**offset)
        if offset > 0:
            prev = get_column_letter(offset + 1)
            cells[f"{letter}3"] = f'=IF(ISNUMBER({prev}2),({letter}2/{prev}2-1)*100,"...")'
    last = get_column_letter(len(years) + 1)
    growth_first = get_column_letter(3)
    cells["A4"] = f"=AVERAGE({growth_first}3:{last}3)"
    workbook = write_workbook(tmp_path / "yoy_avg.xlsx", {"Sheet": cells})
    deflator = series_entry(
        "deflator",
        f"Sheet!B2:{last}2",
        layout="series",
        direction="input",
        header_row=1,
    )
    deflator["exclude_columns"] = ["E"]
    document = bindings_document(
        deflator,
        series_entry(
            "growth",
            f"Sheet!{growth_first}3:{last}3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry("avg", "Sheet!A4", layout="scalar", direction="output"),
        schema_version="1.16.0",
    )
    blanks = ("Sheet!E2",)
    modules = generate_inverted(workbook, document, blank_ranges=blanks)
    internals = modules["internals.py"]
    assert "deflator[time_period]" in internals
    assert "deflator[time_period - 1]" in internals
    catalog, _deps, graph = inverted_graph_parts(workbook, document, blank_ranges=blanks)
    pkg = load_package(modules, tmp_path, name="yoy_avg")
    kwargs = named_input_kwargs(pkg, catalog, graph)
    got_avg = call_compute(pkg, "avg", kwargs)
    got_growth = pkg.internals.growth(deflator=kwargs["deflator"])
    want_avg = _evaluate_cell(workbook, "Sheet!A4", blank_ranges=blanks)
    want_2017 = _evaluate_cell(workbook, "Sheet!E3", blank_ranges=blanks)
    _named_close(got_avg, want_avg, label="YoY AVERAGE")
    _named_close(got_growth[hole_year], want_2017, label="YoY 2017")
    _named_close(got_growth[hole_year], -100.0, label="YoY 2017 vs Excel blank/lag")
