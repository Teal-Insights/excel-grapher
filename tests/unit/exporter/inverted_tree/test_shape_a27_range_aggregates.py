"""Range aggregates over catalog covering series (#667).

Whole-column / whole-row, cross-sheet ranges, `SUM` of a bound series, and
`SUMPRODUCT` fail closed today. Distill each shape as a Tier-1 toy and lower
it with graph-derived access (`covering_series`, same as INDEX/OFFSET).
`xl_sum` / `xl_sumproduct` live in inverted-tree `runtime.py` (core wrappers);
do not embed ctx `export_runtime/`. `SUM(IF(...))` stays fail-closed (#732).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _scalar(value: object) -> object:
    if isinstance(value, tuple):
        assert len(value) == 1
        return value[0]
    return value


def _package_matches_output(
    tmp_path: Path,
    workbook: Path,
    document: dict[str, Any],
    name: str,
    cell: str,
) -> None:
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name=name)
    expected = FormulaEvaluator(graph).evaluate([cell])[cell]
    got = call_compute(pkg, catalog.output_series()[0].series_id, input_kwargs(catalog, graph))
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
    assert pkg.compute_out(src=(1.5, 2.5)) == pytest.approx((4.0,))


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
    assert "take(" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="a27_sum_window")
    assert pkg.compute_out(src=(1.0, 2.0, 100.0)) == pytest.approx((3.0,))
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
    assert pkg.compute_out(left=(1.0, 2.0), right=(3.0, 4.0)) == pytest.approx((11.0,))
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
    assert pkg.compute_out(src=(1.0, 2.0)) == pytest.approx((3.0,))
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
    assert pkg.compute_out(src=(1.0, 2.0)) == pytest.approx((3.0,))
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
    assert pkg.compute_out(left=1.0, right=2.0) == pytest.approx((3.0,))
    _package_matches_output(tmp_path, workbook, document, "a27_cross_eval", "Outputs!Z1")


def test_sum_range_with_unbound_cell_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_unbound.xlsx",
        {
            "Inputs": {"A1": 1.0, "A2": 2.0, "A3": 3.0},
            "Outputs": {"Z1": "=SUM(Inputs!A1:A3)"},
        },
    )
    document = bindings_document(
        series_entry("src", "Inputs!A1:A2", layout="series", direction="input", header_row=10),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    workbook = write_workbook(
        tmp_path / "a27_unbound.xlsx",
        {
            "Inputs": {"A1": 1.0, "A2": 2.0, "A3": 3.0, "A10": 1, "B10": 2},
            "Outputs": {"Z1": "=SUM(Inputs!A1:A3)"},
        },
    )
    with pytest.raises(InvertedTreeExportError, match=r"not a bound|unbound"):
        generate_inverted(workbook, document)


def test_sum_if_range_still_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_sum_if.xlsx",
        {
            "Inputs": {"A1": 1.0, "A2": 2.0, "A10": 1, "B10": 2},
            "Outputs": {"Z1": "=SUM(IF(Inputs!A1:A2>0,Inputs!A1:A2))"},
        },
    )
    document = bindings_document(
        series_entry("src", "Inputs!A1:A2", layout="series", direction="input", header_row=10),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )
    with pytest.raises(
        InvertedTreeExportError,
        match=r"bare range|no inverted-tree runtime helper|unsupported",
    ):
        generate_inverted(workbook, document)
