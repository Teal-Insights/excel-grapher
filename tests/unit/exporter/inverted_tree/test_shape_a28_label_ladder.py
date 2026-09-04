"""Absolute refs into a constant series are static catalog reads, not a lag (#681).

A label-ladder `IF($A$1=$A$8, …, IF($A$1=$A$9, …))` reads the same two cells
from every host member. Those sites are `static` with `coeff = 0`
(`labels[0]` / `labels[1]`), not an aligned lag (`labels[i]` / `labels[i + 1]`).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.access import (
    catalog_index_affine,
    classify_cell_ref_accesses,
)
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    oriented_addresses,
    oriented_document,
    series_entry,
    write_oriented_workbook,
)


def label_ladder_sheets(variant: str = "Low") -> dict[str, dict[str, object]]:
    return {
        "S": {
            "A1": variant,
            "B1": 2008,
            "C1": 2009,
            "A8": "High",
            "A9": "Low",
            "B8": 0,
            "B9": 1,
            "B5": 10,
            "C5": 11,
            "B6": 1,
            "C6": 2,
            "B2": "=IF($A$1=$A$8,B5,IF($A$1=$A$9,B6,0))",
            "C2": "=IF($A$1=$A$8,C5,IF($A$1=$A$9,C6,0))",
        }
    }


def label_ladder_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "variant",
            "S!A1",
            layout="scalar",
            direction="input",
            dtype="string",
        ),
        series_entry(
            "labels",
            "S!A8:A9",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="B",
        ),
        series_entry("hi", "S!B5:C5", layout="series", direction="constant", header_row=1),
        series_entry("lo", "S!B6:C6", layout="series", direction="constant", header_row=1),
        series_entry(
            "picked",
            "S!B2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def label_ladder_workbook(tmp_path: Path, variant: str = "Low") -> Path:
    return write_oriented_workbook(
        tmp_path / f"a28_ladder_{variant.lower()}.xlsx",
        label_ladder_sheets(variant),
        orientation="horizontal",
    )


def _mixed_sheets() -> dict[str, dict[str, object]]:
    return {
        "S": {
            "A1": 0,
            "B1": 2008,
            "C1": 2009,
            "A5": 99,
            "B5": 10,
            "C5": 11,
            "B2": "=IF($A$5=99,B5,0)",
            "C2": "=IF($A$5=99,C5,0)",
        }
    }


def _mixed_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "S!A5:C5", layout="series", direction="constant", header_row=1),
        series_entry(
            "picked",
            "S!B2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def _expected_for(variant: str) -> tuple[float, float]:
    if variant == "High":
        return (10.0, 11.0)
    if variant == "Low":
        return (1.0, 2.0)
    return (0.0, 0.0)


@pytest.mark.parametrize("orientation", ["horizontal", "vertical"])
@pytest.mark.parametrize("variant", ["High", "Low"])
def test_label_ladder_matches_evaluator(
    tmp_path: Path,
    orientation: Literal["horizontal", "vertical"],
    variant: str,
) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a28_{orientation}_{variant.lower()}.xlsx",
        label_ladder_sheets(variant),
        orientation=orientation,
    )
    document = oriented_document(label_ladder_bindings(), orientation)
    pkg = load_package(
        generate_inverted(workbook, document),
        tmp_path,
        name=f"a28_{orientation}_{variant.lower()}",
    )
    cells = list(oriented_addresses(("S!B2", "S!C2"), orientation))
    graph = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph).evaluate(cells)
    got = pkg.compute_picked(variant=variant)
    assert got == pytest.approx((expected[cells[0]], expected[cells[1]]))
    assert got == pytest.approx(_expected_for(variant))


@pytest.mark.parametrize("orientation", ["horizontal", "vertical"])
def test_label_ladder_emits_literal_subscripts(
    tmp_path: Path,
    orientation: Literal["horizontal", "vertical"],
) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a28_emit_{orientation}.xlsx",
        label_ladder_sheets("Low"),
        orientation=orientation,
    )
    document = oriented_document(label_ladder_bindings(), orientation)
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "labels[0]" in internals
    assert "labels[1]" in internals
    assert "labels[i]" not in internals
    assert "labels[i + 1]" not in internals


def test_label_ladder_labels_are_static_not_lagged(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(
        label_ladder_workbook(tmp_path),
        label_ladder_bindings(),
    )
    picked = deps["picked"]
    assert picked.is_scan is False
    assert "labels" not in picked.lagged_ids
    assert "labels" not in picked.aligned_ids
    accesses = classify_cell_ref_accesses(
        catalog.get("picked"),
        catalog.get("labels"),
        catalog,
        graph,
    )
    assert len(accesses) == 2
    affines = sorted(catalog_index_affine(access) for access in accesses)
    assert affines == [(0, 0), (0, 1)]


def test_mixed_absolute_and_relative_are_two_accesses(tmp_path: Path) -> None:
    workbook = write_oriented_workbook(
        tmp_path / "a28_mixed.xlsx",
        _mixed_sheets(),
        orientation="horizontal",
    )
    catalog, deps, graph = inverted_graph_parts(workbook, _mixed_bindings())
    picked = deps["picked"]
    assert "src" not in picked.lagged_ids
    assert "src" not in picked.aligned_ids
    accesses = classify_cell_ref_accesses(
        catalog.get("picked"),
        catalog.get("src"),
        catalog,
        graph,
    )
    assert len(accesses) == 2
    coeffs = sorted(catalog_index_affine(access)[0] for access in accesses)
    assert coeffs == [0, 1]
    pkg = load_package(generate_inverted(workbook, _mixed_bindings()), tmp_path, name="a28_mixed")
    cells = ["S!B2", "S!C2"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_picked()
    assert got == pytest.approx((expected["S!B2"], expected["S!C2"]))
    assert got == pytest.approx((10.0, 11.0))
