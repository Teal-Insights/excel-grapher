"""Range aggregates over catalog covering series (#667, #732, #749).

`SUM` / `SUMPRODUCT` of a bound series, whole-column / whole-row refs, and
cross-sheet ranges lower with graph-derived access (`covering_series`,
`take` for a window). `xl_sum` / `xl_sumproduct` live in inverted-tree
`runtime.py` (core wrappers); do not embed ctx `export_runtime/`.
Array-style `SUM(IF(range,…))` and `SUMPRODUCT(IF(range,…))` lower as
`xl_if` over positional range tables when interiors are element-aligned
(#732); `AVERAGE(IF)` / `MAX(IF)` reuse that emit (#749). Unsound
alignment stays fail-closed. Workbooks that already use `AVERAGEIF` /
`MAXIFS` stay on those functions.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
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


def _scalar(value: object) -> object:
    if isinstance(value, tuple):
        assert len(value) == 1
        return value[0]
    return value


def _output_series_for_cell(catalog: SeriesCatalog, cell: str) -> BoundSeries:
    for series in catalog.output_series():
        if cell in series.cells:
            return series
    raise AssertionError(f"no output series owns {cell}")


def _package_matches_output(
    tmp_path: Path,
    workbook: Path,
    document: dict[str, Any],
    name: str,
    cell: str,
    *,
    pkg: object | None = None,
) -> None:
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    loaded = (
        pkg
        if pkg is not None
        else load_package(generate_inverted(workbook, document), tmp_path, name=name)
    )
    expected = FormulaEvaluator(graph).evaluate([cell])[cell]
    series = _output_series_for_cell(catalog, cell)
    got = call_compute(loaded, series.series_id, named_input_kwargs(loaded, catalog, graph))
    assert _scalar(got) == pytest.approx(expected)


def range_sum_workbook(tmp_path: Path) -> Path:
    """`SUM` of a two-cell bound series (`Inputs!A2:B2`)."""
    return write_workbook(
        tmp_path / "a27_range_sum.xlsx",
        {
            "Inputs": {"A1": 2024, "B1": 2025, "A2": 1.5, "B2": 2.5},
            "Outputs": {"Z1": "=SUM(Inputs!A2:B2)"},
        },
    )


def range_sum_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "Inputs!A2:B2", layout="series", direction="input", header_row=1),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )


def test_sum_of_bound_series_emits_runtime_helper(tmp_path: Path) -> None:
    workbook = range_sum_workbook(tmp_path)
    modules = generate_inverted(workbook, range_sum_bindings())
    assert "xl_sum(" in modules["internals.py"]
    assert "def xl_sum" in modules["runtime.py"]
    pkg = load_package(modules, tmp_path, name="a27_sum_emit")
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(1.5, 2.5))
    ) == pytest.approx(4.0)


def test_sum_of_bound_series_matches_evaluator(tmp_path: Path) -> None:
    workbook = range_sum_workbook(tmp_path)
    _package_matches_output(tmp_path, workbook, range_sum_bindings(), "a27_sum_eval", "Outputs!Z1")


def test_sum_of_series_window_takes_only_the_range(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_window.xlsx",
        {
            "Inputs": {
                "A1": 2024,
                "B1": 2025,
                "C1": 2026,
                "A2": 1.0,
                "B2": 2.0,
                "C2": 100.0,
            },
            "Outputs": {"Z1": "=SUM(Inputs!A2:B2)"},
        },
    )
    document = bindings_document(
        series_entry("src", "Inputs!A2:C2", layout="series", direction="input", header_row=1),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    modules = generate_inverted(workbook, document)
    assert "src[2024]" in modules["internals.py"]
    assert "src[2025]" in modules["internals.py"]
    assert "src[2026]" not in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="a27_sum_window")
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(1.0, 2.0, 100.0))
    ) == pytest.approx(3.0)
    _package_matches_output(tmp_path, workbook, document, "a27_sum_window_eval", "Outputs!Z1")


def test_sumproduct_of_bound_series_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sumproduct.xlsx",
        {
            "Inputs": {
                "A1": 2024,
                "B1": 2025,
                "A2": 1.0,
                "B2": 2.0,
                "A3": 3.0,
                "B3": 4.0,
            },
            "Outputs": {"Z1": "=SUMPRODUCT(Inputs!A2:B2,Inputs!A3:B3)"},
        },
    )
    document = bindings_document(
        series_entry("left", "Inputs!A2:B2", layout="series", direction="input", header_row=1),
        series_entry("right", "Inputs!A3:B3", layout="series", direction="input", header_row=1),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    modules = generate_inverted(workbook, document)
    assert "xl_sumproduct(" in modules["internals.py"]
    assert "def xl_sumproduct" in modules["runtime.py"]
    pkg = load_package(modules, tmp_path, name="a27_sumproduct")
    assert pkg.compute_out(
        left=pkg.data.Left.from_nested(domain=pkg.data.LEFT_DOMAIN, values=(1.0, 2.0)),
        right=pkg.data.Right.from_nested(domain=pkg.data.RIGHT_DOMAIN, values=(3.0, 4.0)),
    ) == pytest.approx(11.0)
    _package_matches_output(tmp_path, workbook, document, "a27_sumproduct_eval", "Outputs!Z1")


def test_sum_whole_column_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_whole_col.xlsx",
        {
            "Inputs": {"A1": 1.0, "A2": 2.0},
            "Outputs": {"Z1": "=SUM(Inputs!A:A)"},
        },
    )
    document = bindings_document(
        series_entry(
            "src",
            "Inputs!A1:A2",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="TIME_PERIOD",
            key_read="int",
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_whole_col")
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(1.0, 2.0))
    ) == pytest.approx(3.0)
    _package_matches_output(tmp_path, workbook, document, "a27_whole_col_eval", "Outputs!Z1")


def test_sum_whole_row_matches_evaluator(tmp_path: Path) -> None:
    document = bindings_document(
        series_entry("src", "Inputs!A1:B1", layout="series", direction="input", header_row=10),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    workbook = write_workbook(
        tmp_path / "a27_whole_row.xlsx",
        {
            "Inputs": {"A1": 1.0, "B1": 2.0, "A10": 1, "B10": 2},
            "Outputs": {"Z1": "=SUM(Inputs!1:1)"},
        },
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_whole_row")
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(1.0, 2.0))
    ) == pytest.approx(3.0)
    _package_matches_output(tmp_path, workbook, document, "a27_whole_row_eval", "Outputs!Z1")


def test_sum_cross_sheet_range_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_cross.xlsx",
        {
            "Inputs": {"A1": 1.0},
            "Other": {"A1": 2.0},
            "Outputs": {"Z1": "=SUM(Inputs!A1:Other!A1)"},
        },
    )
    document = bindings_document(
        series_entry("left", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("right", "Other!A1", layout="scalar", direction="input"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_cross")
    assert pkg.compute_out(left=1.0, right=2.0) == pytest.approx(3.0)
    _package_matches_output(tmp_path, workbook, document, "a27_cross_eval", "Outputs!Z1")


def test_sum_range_with_unbound_cell_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_unbound.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": 1.0,
                "A2": 2.0,
                "A3": 3.0,
            },
            "Outputs": {"Z1": "=SUM(Inputs!A1:A3)"},
        },
    )
    document = bindings_document(
        series_entry("src", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    workbook = write_workbook(
        tmp_path / "a27_unbound.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "A1": 1.0,
                "A2": 2.0,
                "A3": 3.0,
                "A10": 1,
                "B10": 2,
            },
            "Outputs": {"Z1": "=SUM(Inputs!A1:A3)"},
        },
    )
    with pytest.raises(InvertedTreeExportError, match=r"not a bound|unbound"):
        generate_inverted(workbook, document)


def range_sum_if_workbook(tmp_path: Path) -> Path:
    """`SUM(IF(range>0, range))` over a two-cell bound series."""
    return write_workbook(
        tmp_path / "a27_sum_if.xlsx",
        {
            "Inputs": {"Z1": 1, "Z2": 2, "A1": -1.0, "A2": 2.0, "A10": 1, "B10": 2},
            "Outputs": {"Z1": "=SUM(IF(Inputs!A1:A2>0,Inputs!A1:A2))"},
        },
    )


def range_sum_if_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )


def test_sum_if_of_bound_series_emits_runtime_helper(tmp_path: Path) -> None:
    workbook = range_sum_if_workbook(tmp_path)
    modules = generate_inverted(workbook, range_sum_if_bindings())
    assert "xl_if(" in modules["internals.py"]
    assert "def xl_if" in modules["runtime.py"]
    pkg = load_package(modules, tmp_path, name="a27_sum_if_emit")
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(-1.0, 2.0))
    ) == pytest.approx(2.0)
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(1.0, 2.0))
    ) == pytest.approx(3.0)


def test_sum_if_of_bound_series_matches_evaluator(tmp_path: Path) -> None:
    workbook = range_sum_if_workbook(tmp_path)
    _package_matches_output(
        tmp_path, workbook, range_sum_if_bindings(), "a27_sum_if_eval", "Outputs!Z1"
    )


def test_sum_if_then_else_ranges_match_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if_else.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": -1.0,
                "A2": 2.0,
                "B1": 10.0,
                "B2": 20.0,
                "C1": 100.0,
                "C2": 200.0,
                "A10": 1,
                "B10": 2,
                "C10": 3,
            },
            "Outputs": {"Z1": "=SUM(IF(Inputs!A1:A2>0,Inputs!B1:B2,Inputs!C1:C2))"},
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry(
            "then_s", "Inputs!B1:B2", layout="series", direction="input", label_column="Z"
        ),
        series_entry(
            "else_s", "Inputs!C1:C2", layout="series", direction="input", label_column="Z"
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_sum_if_else")
    assert pkg.compute_out(
        flag=pkg.data.Flag.from_nested(domain=pkg.data.FLAG_DOMAIN, values=(-1.0, 2.0)),
        then_s=pkg.data.ThenS.from_nested(domain=pkg.data.THEN_S_DOMAIN, values=(10.0, 20.0)),
        else_s=pkg.data.ElseS.from_nested(domain=pkg.data.ELSE_S_DOMAIN, values=(100.0, 200.0)),
    ) == pytest.approx(120.0)
    _package_matches_output(tmp_path, workbook, document, "a27_sum_if_else_eval", "Outputs!Z1")


def test_sum_if_scalar_else_and_nested_if_match_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if_nested.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": -1.0,
                "A2": 2.0,
                "B1": 10.0,
                "B2": 20.0,
                "C1": 100.0,
                "C2": 200.0,
                "A10": 1,
                "B10": 2,
                "C10": 3,
            },
            "Outputs": {
                "Z1": "=SUM(IF(Inputs!A1:A2>0,Inputs!B1:B2,0))",
                "Z2": "=SUM(IF(Inputs!A1:A2>0,IF(Inputs!B1:B2>15,Inputs!C1:C2,0),0))",
            },
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry(
            "then_s", "Inputs!B1:B2", layout="series", direction="input", label_column="Z"
        ),
        series_entry(
            "else_s", "Inputs!C1:C2", layout="series", direction="input", label_column="Z"
        ),
        series_entry("out_else", "Outputs!Z1", layout="scalar", direction="output"),
        series_entry("out_nest", "Outputs!Z2", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_sum_if_nested")
    assert pkg.compute_out_else(
        flag=pkg.data.Flag.from_nested(domain=pkg.data.FLAG_DOMAIN, values=(-1.0, 2.0)),
        then_s=pkg.data.ThenS.from_nested(domain=pkg.data.THEN_S_DOMAIN, values=(10.0, 20.0)),
    ) == pytest.approx(20.0)
    assert pkg.compute_out_nest(
        flag=pkg.data.Flag.from_nested(domain=pkg.data.FLAG_DOMAIN, values=(-1.0, 2.0)),
        then_s=pkg.data.ThenS.from_nested(domain=pkg.data.THEN_S_DOMAIN, values=(10.0, 20.0)),
        else_s=pkg.data.ElseS.from_nested(domain=pkg.data.ELSE_S_DOMAIN, values=(100.0, 200.0)),
    ) == pytest.approx(200.0)
    _package_matches_output(
        tmp_path, workbook, document, "a27_sum_if_nested_eval", "Outputs!Z1", pkg=pkg
    )
    _package_matches_output(
        tmp_path, workbook, document, "a27_sum_if_nested_eval", "Outputs!Z2", pkg=pkg
    )


def test_sum_if_2d_range_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if_2d.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": 1.0,
                "B1": -2.0,
                "A2": 3.0,
                "B2": 4.0,
                "A10": 1,
                "B10": 2,
            },
            "Outputs": {"Z1": "=SUM(IF(Inputs!A1:B2>0,Inputs!A1:B2,0))"},
        },
    )
    document = bindings_document(
        series_entry("left", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry("right", "Inputs!B1:B2", layout="series", direction="input", label_column="Z"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_sum_if_2d")
    assert pkg.compute_out(
        left=pkg.data.Left.from_nested(domain=pkg.data.LEFT_DOMAIN, values=(1.0, 3.0)),
        right=pkg.data.Right.from_nested(domain=pkg.data.RIGHT_DOMAIN, values=(-2.0, 4.0)),
    ) == pytest.approx(8.0)
    _package_matches_output(tmp_path, workbook, document, "a27_sum_if_2d_eval", "Outputs!Z1")


def test_sum_if_broadcast_scalar_equals_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if_eq.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": -1.0,
                "A2": 2.0,
                "B1": 10.0,
                "B2": 20.0,
                "E1": 2.0,
                "A10": 1,
                "B10": 2,
            },
            "Outputs": {"Z1": "=SUM(IF(Inputs!A1:A2=Inputs!E1,Inputs!B1:B2,0))"},
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry(
            "then_s", "Inputs!B1:B2", layout="series", direction="input", label_column="Z"
        ),
        series_entry("needle", "Inputs!E1", layout="scalar", direction="input"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_sum_if_eq")
    assert pkg.compute_out(
        flag=pkg.data.Flag.from_nested(domain=pkg.data.FLAG_DOMAIN, values=(-1.0, 2.0)),
        then_s=pkg.data.ThenS.from_nested(domain=pkg.data.THEN_S_DOMAIN, values=(10.0, 20.0)),
        needle=2.0,
    ) == pytest.approx(20.0)
    _package_matches_output(tmp_path, workbook, document, "a27_sum_if_eq_eval", "Outputs!Z1")


def test_sum_if_window_of_longer_series_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if_window.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": -1.0,
                "A2": 2.0,
                "A3": 100.0,
                "A10": 1,
                "B10": 2,
                "C10": 3,
            },
            "Outputs": {"Z1": "=SUM(IF(Inputs!A1:A2>0,Inputs!A1:A2))"},
        },
    )
    document = bindings_document(
        series_entry("src", "Inputs!A1:A3", layout="series", direction="input", label_column="Z"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    modules = generate_inverted(workbook, document)
    assert "src[3]" not in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="a27_sum_if_window")
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(-1.0, 2.0, 100.0))
    ) == pytest.approx(2.0)
    _package_matches_output(tmp_path, workbook, document, "a27_sum_if_window_eval", "Outputs!Z1")


def range_sumproduct_if_workbook(tmp_path: Path) -> Path:
    """`SUMPRODUCT(IF(range>0, range))` over a two-cell bound series."""
    return write_workbook(
        tmp_path / "a27_sumproduct_if.xlsx",
        {
            "Inputs": {"Z1": 1, "Z2": 2, "A1": -1.0, "A2": 2.0, "A10": 1, "B10": 2},
            "Outputs": {"Z1": "=SUMPRODUCT(IF(Inputs!A1:A2>0,Inputs!A1:A2))"},
        },
    )


def range_sumproduct_if_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )


def test_sumproduct_if_of_bound_series_emits_runtime_helper(tmp_path: Path) -> None:
    workbook = range_sumproduct_if_workbook(tmp_path)
    modules = generate_inverted(workbook, range_sumproduct_if_bindings())
    assert "xl_if(" in modules["internals.py"]
    assert "xl_sumproduct(" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="a27_sumproduct_if_emit")
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(-1.0, 2.0))
    ) == pytest.approx(2.0)
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(1.0, 2.0))
    ) == pytest.approx(3.0)


def test_sumproduct_if_of_bound_series_matches_evaluator(tmp_path: Path) -> None:
    workbook = range_sumproduct_if_workbook(tmp_path)
    _package_matches_output(
        tmp_path,
        workbook,
        range_sumproduct_if_bindings(),
        "a27_sumproduct_if_eval",
        "Outputs!Z1",
    )


def range_average_if_workbook(tmp_path: Path) -> Path:
    """`AVERAGE(IF(range>0, range))` over a two-cell bound series."""
    return write_workbook(
        tmp_path / "a27_average_if.xlsx",
        {
            "Inputs": {"Z1": 1, "Z2": 2, "A1": -1.0, "A2": 2.0, "A10": 1, "B10": 2},
            "Outputs": {"Z1": "=AVERAGE(IF(Inputs!A1:A2>0,Inputs!A1:A2))"},
        },
    )


def range_average_if_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )


def range_max_if_workbook(tmp_path: Path) -> Path:
    """`MAX(IF(range>0, range))` over a two-cell bound series."""
    return write_workbook(
        tmp_path / "a27_max_if.xlsx",
        {
            "Inputs": {"Z1": 1, "Z2": 2, "A1": -1.0, "A2": 2.0, "A10": 1, "B10": 2},
            "Outputs": {"Z1": "=MAX(IF(Inputs!A1:A2>0,Inputs!A1:A2))"},
        },
    )


def range_max_if_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )


def test_average_if_of_bound_series_emits_runtime_helper(tmp_path: Path) -> None:
    workbook = range_average_if_workbook(tmp_path)
    modules = generate_inverted(workbook, range_average_if_bindings())
    assert "xl_if(" in modules["internals.py"]
    assert "xl_average(" in modules["internals.py"]
    assert "def xl_average" in modules["runtime.py"]
    pkg = load_package(modules, tmp_path, name="a27_average_if_emit")
    # Omitted else is FALSE; AVERAGE skips logicals, so only the matching 2.0.
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(-1.0, 2.0))
    ) == pytest.approx(2.0)
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(1.0, 2.0))
    ) == pytest.approx(1.5)


def test_average_if_of_bound_series_matches_evaluator(tmp_path: Path) -> None:
    workbook = range_average_if_workbook(tmp_path)
    _package_matches_output(
        tmp_path, workbook, range_average_if_bindings(), "a27_average_if_eval", "Outputs!Z1"
    )


def test_max_if_of_bound_series_emits_runtime_helper(tmp_path: Path) -> None:
    workbook = range_max_if_workbook(tmp_path)
    modules = generate_inverted(workbook, range_max_if_bindings())
    assert "xl_if(" in modules["internals.py"]
    assert "xl_max(" in modules["internals.py"]
    assert "def xl_max" in modules["runtime.py"]
    pkg = load_package(modules, tmp_path, name="a27_max_if_emit")
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(-1.0, 2.0))
    ) == pytest.approx(2.0)
    assert pkg.compute_out(
        src=pkg.data.Src.from_nested(domain=pkg.data.SRC_DOMAIN, values=(1.0, 4.0))
    ) == pytest.approx(4.0)


def test_max_if_of_bound_series_matches_evaluator(tmp_path: Path) -> None:
    workbook = range_max_if_workbook(tmp_path)
    _package_matches_output(
        tmp_path, workbook, range_max_if_bindings(), "a27_max_if_eval", "Outputs!Z1"
    )


def test_average_if_then_else_ranges_match_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_average_if_else.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": -1.0,
                "A2": 2.0,
                "B1": 10.0,
                "B2": 20.0,
                "C1": 100.0,
                "C2": 200.0,
                "A10": 1,
                "B10": 2,
                "C10": 3,
            },
            "Outputs": {"Z1": "=AVERAGE(IF(Inputs!A1:A2>0,Inputs!B1:B2,Inputs!C1:C2))"},
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry(
            "then_s", "Inputs!B1:B2", layout="series", direction="input", label_column="Z"
        ),
        series_entry(
            "else_s", "Inputs!C1:C2", layout="series", direction="input", label_column="Z"
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_average_if_else")
    assert pkg.compute_out(
        flag=pkg.data.Flag.from_nested(domain=pkg.data.FLAG_DOMAIN, values=(-1.0, 2.0)),
        then_s=pkg.data.ThenS.from_nested(domain=pkg.data.THEN_S_DOMAIN, values=(10.0, 20.0)),
        else_s=pkg.data.ElseS.from_nested(domain=pkg.data.ELSE_S_DOMAIN, values=(100.0, 200.0)),
    ) == pytest.approx(60.0)
    _package_matches_output(tmp_path, workbook, document, "a27_average_if_else_eval", "Outputs!Z1")


def test_max_if_negative_then_skips_omitted_else(tmp_path: Path) -> None:
    """Omitted else is FALSE, not 0, so a negative match stays the max."""
    workbook = write_workbook(
        tmp_path / "a27_max_if_neg.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": -1.0,
                "A2": 2.0,
                "B1": -10.0,
                "B2": -20.0,
                "A10": 1,
                "B10": 2,
            },
            "Outputs": {"Z1": "=MAX(IF(Inputs!A1:A2>0,Inputs!B1:B2))"},
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry(
            "then_s", "Inputs!B1:B2", layout="series", direction="input", label_column="Z"
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a27_max_if_neg")
    assert pkg.compute_out(
        flag=pkg.data.Flag.from_nested(domain=pkg.data.FLAG_DOMAIN, values=(-1.0, 2.0)),
        then_s=pkg.data.ThenS.from_nested(domain=pkg.data.THEN_S_DOMAIN, values=(-10.0, -20.0)),
    ) == pytest.approx(-20.0)
    _package_matches_output(tmp_path, workbook, document, "a27_max_if_neg_eval", "Outputs!Z1")


def test_averageif_is_not_rewritten_as_array_if(tmp_path: Path) -> None:
    """Native `AVERAGEIF` stays that function; it is not rewritten as `AVERAGE(IF)`."""
    workbook = write_workbook(
        tmp_path / "a27_averageif.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": -1.0,
                "A2": 2.0,
                "A10": 1,
                "B10": 2,
            },
            "Outputs": {"Z1": '=AVERAGEIF(Inputs!A1:A2,">0")'},
        },
    )
    document = bindings_document(
        series_entry("src", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    with pytest.raises(
        InvertedTreeExportError,
        match=r"bare range in value position|no inverted-tree runtime helper",
    ):
        generate_inverted(workbook, document)


@pytest.mark.parametrize("outer", ["AVERAGE", "MAX"])
def test_average_and_max_if_unsound_alignment_fails_closed(tmp_path: Path, outer: str) -> None:
    workbook = write_workbook(
        tmp_path / "a27_avg_max_if_closed.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": 1.0,
                "A2": 2.0,
                "B1": 10.0,
                "B2": 20.0,
                "B3": 30.0,
                "B4": 40.0,
                "B5": 50.0,
                "A10": 1,
                "B10": 2,
            },
            "Outputs": {"Z1": f"={outer}(IF(Inputs!A1:A2>0,Inputs!B1:B5,0))"},
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry(
            "then_s", "Inputs!B1:B5", layout="series", direction="input", label_column="Z"
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match=r"array IF shape mismatch"):
        generate_inverted(workbook, document)


@pytest.mark.parametrize(
    ("formula", "match"),
    [
        (
            "=SUM(IF(Inputs!A1:A2>0,Inputs!B1:B5,0))",
            r"array IF shape mismatch",
        ),
        (
            "=SUM(IF(Inputs!A1:A2>0,SUM(Inputs!B1:B2),0))",
            r"array IF nested aggregate is unsupported",
        ),
        (
            "=SUM(IF(AND(Inputs!A1:A2>0,Inputs!E1>0),Inputs!B1:B2,0))",
            r"array IF AND/OR collapse is unsupported",
        ),
        (
            "=SUM(IF(Inputs!A1:A2>0,-Inputs!B1:B2,0))",
            r"array IF unary '-' is unsupported",
        ),
        (
            "=SUM(IF(Inputs!A1:A2>0,ABS(Inputs!B1:B2),0))",
            r"array IF interior ABS is unsupported",
        ),
        (
            '=SUM(IF(Inputs!A1:A2&"x"="1x",Inputs!B1:B2,0))',
            r"array IF operator '&' is unsupported",
        ),
        (
            "=SUM(IF(IFS(Inputs!A1:A2>0,TRUE),Inputs!B1:B2,0))",
            r"array IF IFS is unsupported",
        ),
        (
            "=SUM(IF(Inputs!A1:A2>0,CHOOSE(1,Inputs!B1:B2),0))",
            r"array IF CHOOSE is unsupported",
        ),
        (
            "=SUM(IF(Inputs!A1:A2>0,SWITCH(1,1,Inputs!B1:B2),0))",
            r"array IF SWITCH is unsupported",
        ),
    ],
)
def test_sum_if_unsound_alignment_fails_closed(tmp_path: Path, formula: str, match: str) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if_closed.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": 1.0,
                "A2": 2.0,
                "B1": 10.0,
                "B2": 20.0,
                "B3": 30.0,
                "B4": 40.0,
                "B5": 50.0,
                "E1": 1.0,
                "A10": 1,
                "B10": 2,
            },
            "Outputs": {"Z1": formula},
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry(
            "then_s", "Inputs!B1:B5", layout="series", direction="input", label_column="Z"
        ),
        series_entry("extra", "Inputs!E1", layout="scalar", direction="input"),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match=match):
        generate_inverted(workbook, document)


def test_sum_if_whole_column_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if_whole_col.xlsx",
        {
            "Inputs": {"A1": 1.0, "A2": 2.0, "B1": 10.0, "B2": 20.0},
            "Outputs": {"Z1": "=SUM(IF(Inputs!A:A>0,Inputs!B:B,0))"},
        },
    )
    document = bindings_document(
        series_entry(
            "flag",
            "Inputs!A1:A2",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="TIME_PERIOD",
            key_read="int",
        ),
        series_entry(
            "then_s",
            "Inputs!B1:B2",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="TIME_PERIOD",
            key_read="int",
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match=r"array IF does not support whole-column"):
        generate_inverted(workbook, document)


def test_sum_if_at_operator_has_no_formula_ast(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if_at.xlsx",
        {
            "Inputs": {
                "Z1": 1,
                "Z2": 2,
                "Z3": 3,
                "Z4": 4,
                "Z5": 5,
                "A1": 1.0,
                "A2": 2.0,
                "B1": 10.0,
                "B2": 20.0,
                "A10": 1,
                "B10": 2,
            },
            "Outputs": {"Z1": "=SUM(IF(@Inputs!A1:A2>0,Inputs!B1:B2,0))"},
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1:A2", layout="series", direction="input", label_column="Z"),
        series_entry(
            "then_s", "Inputs!B1:B2", layout="series", direction="input", label_column="Z"
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match=r"no formula AST"):
        generate_inverted(workbook, document)
