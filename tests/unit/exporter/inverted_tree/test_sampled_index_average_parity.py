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

from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
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
from tests.unit.exporter.inverted_tree.test_off_axis_producer_keys import (
    _ALPHA,
    _BUCKET_ORIGIN,
    _N_HOST,
    _ORIGIN,
    _letter,
    _matrix_copy_bindings,
    _matrix_copy_workbook,
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
    """`=producer+0` off-axis members must lower to 0 so AVERAGE includes them."""
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
    """Unsampled interior off-axis `+0` copy must not drop out of AVERAGE."""
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
    got = pkg.compute_schedule(principal=pkg.data.PRINCIPAL_DEFAULT)
    tail_year = _ORIGIN + n_host - 1
    tail_cell = f"Host!{_letter(n_host - 1)}2"
    want = _evaluate_cell(workbook, tail_cell, blank_ranges=blanks)
    _named_close(got[_ALPHA, tail_year], want, label=f"Alpha {tail_year}")


def test_yoy_average_keeps_lag_year_before_series_origin(tmp_path: Path) -> None:
    """A1-like AVERAGE of YoY growth must include the first year (issue 892).

    Deflator is bound from 2014, but 2013 exists as a valued leaf the first YoY
    formula reads. Folding `deflator[t]/deflator[t-1]` must not blank 2014.
    """
    years = tuple(range(2013, 2024))
    cells: dict[str, object] = {}
    for offset, year in enumerate(years):
        letter = get_column_letter(offset + 2)
        cells[f"{letter}1"] = year
        cells[f"{letter}2"] = 100.0 * (1.03**offset)
        if offset > 0:
            prev = get_column_letter(offset + 1)
            cells[f"{letter}3"] = f'=IF(ISNUMBER({prev}2),({letter}2/{prev}2-1)*100,"...")'
    last = get_column_letter(len(years) + 1)
    second = get_column_letter(3)
    cells["A4"] = f"=AVERAGE({second}3:{last}3)"
    workbook = write_workbook(tmp_path / "yoy_avg.xlsx", {"Sheet": cells})
    document = bindings_document(
        series_entry("seed", "Sheet!B2", layout="scalar", direction="input"),
        series_entry(
            "deflator",
            f"Sheet!C2:{last}2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "growth",
            f"Sheet!C3:{last}3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry("avg", "Sheet!A4", layout="scalar", direction="output"),
        schema_version="1.16.0",
    )
    got = _exported_avg(workbook, document, tmp_path, "yoy_avg", blank_ranges=())
    want = _evaluate_cell(workbook, "Sheet!A4", blank_ranges=())
    _named_close(got, want, label="YoY AVERAGE")
