"""Design properties from `plans/inverted-tree-scheduling.md` (§7–§8).

These replace emission-syntax greps: a corpus-wide differential oracle
(auto rung vs forced fused vs forced rung 3 vs `FormulaEvaluator`),
Θ(S+E) size invariance, orientation via the transpose helper,
schedule-order safety for lexically mis-ordered keys, and
`inspect.signature` leaf-closure contracts.
"""

from __future__ import annotations

import inspect
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog
from excel_grapher.exporter.inverted_tree.schedule import plan_fused_scc, plan_scc
from excel_grapher.grapher.graph import DependencyGraph
from tests.unit.exporter.inverted_tree import test_shape_a20_matrix_join as a20
from tests.unit.exporter.inverted_tree import test_shape_a22_guarded_residual as a22_guarded
from tests.unit.exporter.inverted_tree import test_shape_a22_shift_k as a22_shift_k
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
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
    write_workbook,
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
from tests.unit.exporter.inverted_tree.test_shape_a26_index_block import (
    _country_table_bindings,
    _country_table_workbook,
    country_table_bindings,
    country_table_sheets,
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


def _output_cells_in_export_order(series: BoundSeries) -> tuple[str, ...]:
    """Cells in the order `compute_*` returns values.

    1-D `layout: series` helpers follow catalog expansion order (lexical-key
    safety: `Y9` then `Y10`). A `layout: matrix` series is a key-field nest —
    leading keys outer, `TIME_PERIOD` inner — which matches
    `sorted(domain, key=key_fields)` and diverges from sheet order when the
    matrix is transposed.
    """
    if series.layout != "matrix" or not series.key_fields:
        return series.cells
    keyed = [
        (tuple(point[field] for field in series.key_fields), cell)
        for point, cell in zip(series.domain, series.cells, strict=True)
    ]
    try:
        keyed.sort(key=lambda item: item[0])
    except TypeError:
        return series.cells
    return tuple(cell for _key, cell in keyed)


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
        want = tuple(expected[cell] for cell in _output_cells_in_export_order(series))
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


def _a22_shift_k_workbook(tmp_path: Path) -> Path:
    return a22_shift_k._stride_k_workbook(tmp_path, 5, 2, stem="a22_orient_k")


def _a22_shift_k_bindings() -> dict[str, Any]:
    return a22_shift_k._stride_k_bindings(5)


def _compare_workbook(tmp_path: Path) -> Path:
    """Type-rank comparison + arithmetic coercion (#651)."""
    return write_workbook(
        tmp_path / "compare_rank.xlsx",
        {
            "Engine": {
                "A1": 1,
                "B1": 2,
                "C1": 3,
                "D1": 4,
                "E1": 5,
                "F1": 6,
                "G1": 7,
                "A2": True,
                "B2": "10",
                "C2": "",
                "D2": "abc",
                "E2": "a",
                "F2": "10",
                "G2": True,
                "A3": 100,
                "B3": 10,
                "C3": 0,
                "D3": "ABC",
                "E3": 1,
                "F3": 2,
                "G3": 1,
                "A4": "=IF(A2>A3,1,0)",
                "B4": "=IF(B2=B3,1,0)",
                "C4": "=IF(C2=C3,1,0)",
                "D4": "=IF(D2=D3,1,0)",
                "E4": "=IF(E2<E3,1,0)",
                "F4": "=F2+F3",
                "G4": "=IF(G2=G3,1,0)",
            }
        },
    )


def _compare_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("left", "Engine!A2:G2", layout="series", direction="input", header_row=1),
        series_entry("right", "Engine!A3:G3", layout="series", direction="input", header_row=1),
        series_entry("result", "Engine!A4:G4", layout="series", direction="output", header_row=1),
    )


# Cases a property test cannot yet pass must use `pytest.mark.xfail(strict=True)`,
# not a silent filter (#633, #640). `_CORPUS` is the full oracle.
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
    ("a20_elementwise", a20._elementwise_workbook, a20._elementwise_bindings),
    ("a20_zipper", a20._zipper_workbook, a20._zipper_bindings),
    ("a22_guarded", a22_guarded._series_may_cycle_workbook, a22_guarded._series_may_cycle_bindings),
    ("a22_shift_k", _a22_shift_k_workbook, _a22_shift_k_bindings),
    ("country_table", _country_table_workbook, _country_table_bindings),
    ("compare_rank", _compare_workbook, _compare_bindings),
]


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
    _CORPUS,
    ids=[item[0] for item in _CORPUS],
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


@pytest.mark.parametrize(
    ("case", "workbook_fn", "bindings_fn"),
    _CORPUS,
    ids=[item[0] for item in _CORPUS],
)
def test_corpus_rung2_matches_evaluator_and_auto(
    tmp_path: Path,
    case: str,
    workbook_fn: Callable[[Path], Path],
    bindings_fn: Callable[[], dict[str, Any]],
) -> None:
    workbook = workbook_fn(tmp_path)
    document = bindings_fn()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    auto = load_package(generate_inverted(workbook, document), tmp_path, name=f"{case}_r2_auto")
    forced = load_package(
        generate_inverted(workbook, document, force_rung=2),
        tmp_path,
        name=f"{case}_r2",
    )
    _package_matches_evaluator(auto, catalog, graph)
    _package_matches_evaluator(forced, catalog, graph)
    kwargs = input_kwargs(catalog, graph)
    for series in catalog.output_series():
        _values_close(
            call_compute(auto, series.series_id, kwargs),
            call_compute(forced, series.series_id, kwargs),
        )


def test_force_rung_2_falls_through_when_not_fusible(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "mixed_rung2.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+C3",
                "C2": "=B2+10",
                "A3": "=B2*0.1",
                "B3": "=A2*0.1",
                "C3": "=10",
            },
        },
    )
    document = bindings_document(
        series_entry("x", "Engine!A2:C2", layout="series", direction="output", header_row=1),
        series_entry("y", "Engine!A3:C3", layout="series", direction="internal", header_row=1),
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    assert plan_fused_scc(("x", "y"), catalog=catalog, graph=graph) is None
    assert plan_scc(("x", "y"), catalog=catalog, graph=graph).rung == 3
    modules = generate_inverted(workbook, document, force_rung=2)
    pkg = load_package(modules, tmp_path, name="mixed_rung2")
    _package_matches_evaluator(pkg, catalog, graph)


_ORIENTABLE = [
    ("a10_lag", _lag_workbook, _lag_bindings),
    ("a11_zipper", _zipper_workbook, _zipper_bindings),
    ("a12_shapes", _a12_workbook, _a12_bindings),
    ("a13_flip", _two_series_workbook, _two_series_bindings),
    ("a17_overlap", _overlap_workbook, _overlap_bindings),
    ("a19_terminal", _horizontal_terminal_workbook, _horizontal_terminal_bindings),
    ("a20_elementwise", a20._elementwise_workbook, a20._elementwise_bindings),
    ("a20_zipper", a20._zipper_workbook, a20._zipper_bindings),
    ("a22_guarded", a22_guarded._series_may_cycle_workbook, a22_guarded._series_may_cycle_bindings),
    ("a22_shift_k", _a22_shift_k_workbook, _a22_shift_k_bindings),
    ("country_table", _country_table_workbook, _country_table_bindings),
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


def _constant_label_sheets(n: int) -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {}
    for index in range(n):
        col = get_column_letter(index + 1)
        cells[f"{col}1"] = 2009 + index
        cells[f"{col}2"] = index + 1
        cells[f"{col}3"] = 10.0
        cells[f"{col}4"] = f"={col}3+IF({col}2>=1,1,0)"
    return {"Engine": cells}


def _constant_label_bindings(n: int) -> dict[str, Any]:
    last = get_column_letter(n)
    return bindings_document(
        series_entry(
            "labels",
            f"Engine!A2:{last}2",
            layout="series",
            direction="constant",
            dtype="int",
            header_row=1,
        ),
        series_entry(
            "values",
            f"Engine!A3:{last}3",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "result",
            f"Engine!A4:{last}4",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


_SIZE_CASES: list[
    tuple[str, Callable[[int], Mapping[str, Mapping[str, object]]], Callable[[int], dict[str, Any]]]
] = [
    ("scan", _scan_sheets, lambda n: _series_bindings("path", n)),
    ("zipper", _zipper_sheets, _zipper_bindings_n),
    ("backward", _backward_sheets, lambda n: _series_bindings("value", n)),
    ("country_table", country_table_sheets, country_table_bindings),
    ("constants", _constant_label_sheets, _constant_label_bindings),
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
        assert abs(max(map(len, small_lines)) - max(map(len, large_lines))) <= 8


def test_code_size_independent_of_constant_series_count(tmp_path: Path) -> None:
    """`api.py` does not grow with unused catalog constants or used-constant arity."""

    def generate_with_unused(count: int) -> dict[str, str]:
        sub = tmp_path / f"unused_{count}"
        sub.mkdir()
        workbook = _a1_workbook(sub)
        from fastpyxl import load_workbook

        book = load_workbook(workbook)
        for index in range(count):
            book["Inputs"][f"Z{index + 1}"] = 0.0
        book.save(workbook)
        extras = [
            series_entry(
                f"unused_const_{index}",
                f"Inputs!Z{index + 1}",
                layout="scalar",
                direction="constant",
            )
            for index in range(count)
        ]
        return generate_inverted(workbook, bindings_document(*_a1_bindings()["series"], *extras))

    small_unused = generate_with_unused(2)
    large_unused = generate_with_unused(12)
    assert small_unused["api.py"] == large_unused["api.py"]

    def generate_with_used(count: int) -> tuple[dict[str, str], ModuleType]:
        inputs: dict[str, object] = {"A1": 1.0}
        terms = ["Inputs!A1"]
        extras: list[dict[str, Any]] = []
        for index in range(count):
            cell = f"B{index + 1}"
            inputs[cell] = float(index + 1)
            terms.append(f"Inputs!{cell}")
            extras.append(
                series_entry(
                    f"const_{index}",
                    f"Inputs!{cell}",
                    layout="scalar",
                    direction="constant",
                )
            )
        workbook = write_workbook(
            tmp_path / f"used_{count}.xlsx",
            {"Inputs": inputs, "Outputs": {"A1": "=" + "+".join(terms)}},
        )
        document = bindings_document(
            series_entry("value", "Inputs!A1", layout="scalar", direction="input"),
            *extras,
            series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
        )
        modules = generate_inverted(workbook, document)
        pkg = load_package(modules, tmp_path, name=f"used_c{count}")
        return modules, pkg

    small_used, small_pkg = generate_with_used(2)
    large_used, large_pkg = generate_with_used(8)
    assert all_param_names(small_pkg.compute_result) == ("value",)
    assert all_param_names(large_pkg.compute_result) == ("value",)
    assert small_pkg.compute_result.__constants__ == ("const_0", "const_1")
    assert large_pkg.compute_result.__constants__ == tuple(f"const_{i}" for i in range(8))
    assert "require_length" not in small_used["api.py"]
    assert "require_length" not in large_used["api.py"]
    assert "from .data import" not in large_used["api.py"]
    small_sig_lines = [
        line for line in small_used["api.py"].splitlines() if line.startswith("    ")
    ]
    large_sig_lines = [
        line for line in large_used["api.py"].splitlines() if line.startswith("    ")
    ]
    small_params = [line for line in small_sig_lines if line.strip().startswith("value:")]
    large_params = [line for line in large_sig_lines if line.strip().startswith("value:")]
    assert small_params == large_params


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
    assert pkg.compute_output_path.__constants__ == ()
    assert pkg.compute_output_year1.__constants__ == ()
    public = [name for name in dir(pkg) if inspect.isfunction(getattr(pkg, name, None))]
    assert not any(name.startswith("set_") for name in public)
    assert "make_context" not in public

    pkg5 = load_package(
        generate_inverted(_a5_workbook(tmp_path), _a5_bindings()), tmp_path, name="sig_a5"
    )
    shocked_sig = inspect.signature(pkg5.compute_output_shocked)
    assert set(required_param_names(pkg5.compute_output_shocked)) == {"value", "shock_year"}
    assert "engine_year_labels" not in shocked_sig.parameters
    assert "engine_year_labels" not in inspect.signature(pkg5.compute_output_baseline).parameters
    assert pkg5.compute_output_shocked.__constants__ == ("engine_year_labels",)
    assert pkg5.compute_output_baseline.__constants__ == ()
