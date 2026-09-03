"""Design properties from `plans/inverted-tree-scheduling.md` (§7–§8).

These replace emission-syntax greps: a corpus-wide differential oracle
(auto rung vs forced rung 3 vs `FormulaEvaluator`), Θ(S+E) size
invariance, orientation via the transpose helper, schedule-order safety
for lexically mis-ordered keys, and `inspect.signature` leaf-closure
contracts.
"""

from __future__ import annotations

import inspect
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import SeriesCatalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import plan_fused_scc
from excel_grapher.grapher.graph import DependencyGraph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    oriented_document,
    required_param_names,
    series_entry,
    transpose_sheets,
    write_oriented_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import (
    _a1_bindings,
    _a1_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a5_constants import (
    _a5_bindings,
    _a5_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a10_other_series_lag import (
    _lag_bindings,
    _lag_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _zipper_bindings,
    _zipper_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a12_formula_shape import (
    _a12_bindings,
    _a12_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a13_identity_flip import (
    _two_series_bindings,
    _two_series_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a16_exp import (
    _exp_bindings,
    _exp_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a17_overlap_take import (
    _overlap_bindings,
    _overlap_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a19_demand_floor import (
    _horizontal_terminal_bindings,
    _horizontal_terminal_workbook,
    _stride2_terminal_bindings,
    _stride2_terminal_workbook,
)


def _values_close(got: object, expected: object) -> None:
    if isinstance(got, tuple) and isinstance(expected, tuple):
        assert len(got) == len(expected)
        for left, right in zip(got, expected, strict=True):
            _values_close(left, right)
        return
    if isinstance(got, str) or isinstance(expected, str):
        assert got == expected
        return
    assert got == pytest.approx(expected)


def _package_matches_evaluator(
    pkg: object,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> None:
    kwargs = input_kwargs(catalog, graph)
    cells = [cell for series in catalog.output_series() for cell in series.cells]
    expected = FormulaEvaluator(graph).evaluate(cells)
    for series in catalog.output_series():
        name = series.compute_name or f"compute_{series.series_id}"
        function = getattr(pkg, name)
        accepted = set(inspect.signature(function).parameters)
        got = function(**{key: value for key, value in kwargs.items() if key in accepted})
        if not isinstance(got, tuple):
            got = (got,)
        want = tuple(expected[cell] for cell in series.cells)
        _values_close(got, want)


def _emit_and_compare(
    workbook: Path,
    document: dict[str, Any],
    tmp_path: Path,
    name: str,
    *,
    force_rung: int | None = None,
) -> None:
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    modules = generate_inverted(workbook, document, force_rung=force_rung)
    pkg = load_package(modules, tmp_path, name=name)
    _package_matches_evaluator(pkg, catalog, graph)


_CORPUS: list[tuple[str, Callable[[Path], Path], Callable[[], dict[str, Any]]]] = [
    ("a1_leaf", _a1_workbook, _a1_bindings),
    ("a5_constants", _a5_workbook, _a5_bindings),
    ("a10_lag", _lag_workbook, _lag_bindings),
    ("a11_zipper", _zipper_workbook, _zipper_bindings),
    ("a12_shapes", _a12_workbook, _a12_bindings),
    ("a13_flip", _two_series_workbook, _two_series_bindings),
    ("a17_overlap", _overlap_workbook, _overlap_bindings),
    ("a19_terminal", _horizontal_terminal_workbook, _horizontal_terminal_bindings),
    ("a19_stride2", _stride2_terminal_workbook, _stride2_terminal_bindings),
]

# Overlapping `take` windows still emit catalog-relative indexes in rung-3
# helpers (#626); the auto rung matches the evaluator for those shapes.
_RUNG3_CORPUS = [item for item in _CORPUS if item[0] != "a17_overlap"]


@pytest.mark.parametrize(
    ("case", "workbook_fn", "bindings_fn"),
    _CORPUS,
    ids=[item[0] for item in _CORPUS],
)
def test_corpus_auto_rung_matches_evaluator(
    tmp_path: Path,
    case: str,
    workbook_fn: Callable[[Path], Path],
    bindings_fn: Callable[[], dict[str, Any]],
) -> None:
    _emit_and_compare(workbook_fn(tmp_path), bindings_fn(), tmp_path, f"{case}_auto")


@pytest.mark.parametrize(
    ("case", "workbook_fn", "bindings_fn"),
    _RUNG3_CORPUS,
    ids=[item[0] for item in _RUNG3_CORPUS],
)
def test_corpus_rung3_matches_evaluator_and_auto(
    tmp_path: Path,
    case: str,
    workbook_fn: Callable[[Path], Path],
    bindings_fn: Callable[[], dict[str, Any]],
) -> None:
    workbook = workbook_fn(tmp_path)
    document = bindings_fn()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    auto = load_package(generate_inverted(workbook, document), tmp_path, name=f"{case}_r3_auto")
    forced = load_package(
        generate_inverted(workbook, document, force_rung=3),
        tmp_path,
        name=f"{case}_r3",
    )
    _package_matches_evaluator(auto, catalog, graph)
    _package_matches_evaluator(forced, catalog, graph)
    kwargs = input_kwargs(catalog, graph)
    for series in catalog.output_series():
        _values_close(
            call_compute(auto, series.series_id, kwargs),
            call_compute(forced, series.series_id, kwargs),
        )


def test_exp_rung3_matches_evaluator(tmp_path: Path) -> None:
    workbook = _exp_workbook(tmp_path, "=EXP(A1)", x=1)
    _emit_and_compare(workbook, _exp_bindings(), tmp_path, "a16_r3", force_rung=3)


def test_force_rung_2_raises_on_mixed_direction_scc(tmp_path: Path) -> None:
    workbook = write_oriented_workbook(
        tmp_path / "mixed_rung2.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "A2": "=B4",
                "B2": "=A4",
                "A4": "=1",
                "B4": "=A2",
            }
        },
        orientation="horizontal",
    )
    document = _two_series_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    assert plan_fused_scc(("x", "y"), catalog=catalog, graph=graph) is None
    with pytest.raises(InvertedTreeExportError, match="force_rung=2"):
        generate_inverted(workbook, document, force_rung=2)


_ORIENTABLE = [
    ("a10_lag", _lag_workbook, _lag_bindings),
    ("a11_zipper", _zipper_workbook, _zipper_bindings),
    ("a12_shapes", _a12_workbook, _a12_bindings),
    ("a13_flip", _two_series_workbook, _two_series_bindings),
    ("a17_overlap", _overlap_workbook, _overlap_bindings),
    ("a19_terminal", _horizontal_terminal_workbook, _horizontal_terminal_bindings),
]


@pytest.mark.parametrize(
    ("case", "workbook_fn", "bindings_fn"),
    _ORIENTABLE,
    ids=[item[0] for item in _ORIENTABLE],
)
def test_corpus_vertical_orientation_matches_evaluator(
    tmp_path: Path,
    case: str,
    workbook_fn: Callable[[Path], Path],
    bindings_fn: Callable[[], dict[str, Any]],
) -> None:
    from tests.unit.exporter.inverted_tree.helpers import load_workbook_sheets

    source = workbook_fn(tmp_path)
    sheets = transpose_sheets(load_workbook_sheets(source))
    workbook = write_oriented_workbook(
        tmp_path / f"{case}_vertical.xlsx", sheets, orientation="horizontal"
    )
    document = oriented_document(bindings_fn(), "vertical")
    _emit_and_compare(workbook, document, tmp_path, f"{case}_vert")


def _scan_sheets(n: int) -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {}
    for index in range(n):
        col = get_column_letter(index + 1)
        cells[f"{col}1"] = 2009 + index
        if index == 0:
            cells[f"{col}2"] = "=100"
        else:
            pred = get_column_letter(index)
            cells[f"{col}2"] = f"={pred}2*1.02"
    return {"Engine": cells}


def _zipper_sheets(n: int) -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {}
    for index in range(n):
        col = get_column_letter(index + 1)
        cells[f"{col}1"] = 2009 + index
        if index == 0:
            cells[f"{col}2"] = "=100"
        else:
            pred = get_column_letter(index)
            cells[f"{col}2"] = f"={pred}2+{col}3"
            cells[f"{col}3"] = f"={pred}2*0.02"
    return {"Engine": cells}


def _backward_sheets(n: int) -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {}
    for index in range(n):
        col = get_column_letter(index + 1)
        cells[f"{col}1"] = 2009 + index
        if index + 1 == n:
            cells[f"{col}2"] = "=100"
        else:
            nxt = get_column_letter(index + 2)
            cells[f"{col}2"] = f"={nxt}2*0.9"
    return {"Engine": cells}


def _series_bindings(series_id: str, n: int, extra: Sequence[dict[str, Any]] = ()) -> dict:
    last = get_column_letter(n)
    entries = [
        series_entry(
            series_id,
            f"Engine!A2:{last}2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        *extra,
    ]
    return bindings_document(*entries)


def _zipper_bindings_n(n: int) -> dict[str, Any]:
    last = get_column_letter(n)
    return bindings_document(
        series_entry(
            "debt",
            f"Engine!A2:{last}2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        series_entry(
            "adjustment",
            f"Engine!B3:{last}3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


_SIZE_CASES: list[
    tuple[str, Callable[[int], Mapping[str, Mapping[str, object]]], Callable[[int], dict[str, Any]]]
] = [
    ("scan", _scan_sheets, lambda n: _series_bindings("path", n)),
    ("zipper", _zipper_sheets, _zipper_bindings_n),
    ("backward", _backward_sheets, lambda n: _series_bindings("value", n)),
]


@pytest.mark.parametrize(
    ("case", "sheets_fn", "bindings_fn"),
    _SIZE_CASES,
    ids=[item[0] for item in _SIZE_CASES],
)
def test_code_size_independent_of_period_count(
    tmp_path: Path,
    case: str,
    sheets_fn: Callable[[int], Mapping[str, Mapping[str, object]]],
    bindings_fn: Callable[[int], dict[str, Any]],
) -> None:
    small_n, large_n = 6, 24
    small_wb = write_oriented_workbook(
        tmp_path / f"{case}_n6.xlsx", sheets_fn(small_n), orientation="horizontal"
    )
    large_wb = write_oriented_workbook(
        tmp_path / f"{case}_n24.xlsx", sheets_fn(large_n), orientation="horizontal"
    )
    small = generate_inverted(small_wb, bindings_fn(small_n))
    large = generate_inverted(large_wb, bindings_fn(large_n))
    for filename in ("api.py", "internals.py"):
        small_lines = small[filename].splitlines()
        large_lines = large[filename].splitlines()
        assert len(small_lines) == len(large_lines), (
            f"{filename} grew from {len(small_lines)} to {len(large_lines)} lines"
        )
        assert abs(len(small[filename]) - len(large[filename])) <= 48


def test_lexically_misordered_string_keys_match_evaluator(tmp_path: Path) -> None:
    """`Y9 < Y10` lexically, but the catalog order is still Y9, Y10, Y11."""
    sheets = {
        "Engine": {
            "A1": "Y9",
            "B1": "Y10",
            "C1": "Y11",
            "A2": "=100",
            "B2": "=A2+1",
            "C2": "=B2+1",
        }
    }
    workbook = write_oriented_workbook(tmp_path / "lex_keys.xlsx", sheets, orientation="horizontal")
    document = bindings_document(
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
            key_read="string",
        )
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    for force_rung in (None, 3):
        modules = generate_inverted(workbook, document, force_rung=force_rung)
        pkg = load_package(modules, tmp_path, name=f"lex_{force_rung}")
        _package_matches_evaluator(pkg, catalog, graph)
        assert pkg.compute_path() == pytest.approx((100.0, 101.0, 102.0))


def test_leaf_closure_signatures_use_inspect(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_a1_workbook(tmp_path), _a1_bindings()), tmp_path, name="sig_a1"
    )
    path_sig = inspect.signature(pkg.compute_output_path)
    year_sig = inspect.signature(pkg.compute_output_year1)
    assert set(required_param_names(pkg.compute_output_path)) == {
        "initial_debt",
        "growth",
        "interest",
    }
    assert "unused_flag" not in path_sig.parameters
    assert "ctx" not in path_sig.parameters
    assert "unused_flag" not in year_sig.parameters
    assert "ctx" not in year_sig.parameters
    public = [name for name in dir(pkg) if inspect.isfunction(getattr(pkg, name, None))]
    assert not any(name.startswith("set_") for name in public)
    assert "make_context" not in public
