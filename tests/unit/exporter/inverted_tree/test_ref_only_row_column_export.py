"""Address-only ROW/COLUMN/ROWS/COLUMNS refs must not require catalog ownership (#987)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    invoke_public_compute,
    load_package,
    series_entry,
    write_workbook,
)


def _graph_keys(workbook: Path, *targets: str) -> set[str]:
    graph = create_dependency_graph(workbook, list(targets))
    return set(graph.leaf_keys()) | set(graph.formula_keys())


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=ROW($B$1)", 1.0),
        ("=COLUMN($B$1)", 2.0),
        ("=ROWS($B$1:$B$2)", 2.0),
        ("=COLUMNS($B$1:$C$1)", 2.0),
        ("=ROW($B$1)+COLUMN($C$2)", 4.0),
    ],
)
def test_address_only_geometry_exports_without_the_referenced_cells(
    tmp_path: Path, formula: str, expected: float
) -> None:
    workbook = write_workbook(
        tmp_path / "ref_only.xlsx",
        {"Sheet1": {"B1": 99, "B2": 98, "C1": 97, "C2": 96, "A1": formula}},
    )
    assert "Sheet1!B1" not in _graph_keys(workbook, "Sheet1!A1")
    document = bindings_document(series_entry("result", "Sheet1!A1", direction="output"))
    package = load_package(generate_inverted(workbook, document), tmp_path)
    assert invoke_public_compute(package, package.compute_result, {}) == pytest.approx(expected)


def test_offset_index_row_anchor_ignores_off_graph_cell(tmp_path: Path) -> None:
    """LIC-DSF shape: ROW($B$1) is an address anchor, not a series read."""
    workbook = write_workbook(
        tmp_path / "offset_anchor.xlsx",
        {
            "Sheet1": {
                "A1": "k1",
                "A2": "k2",
                "B1": 99,
                "C1": 10,
                "C2": 20,
                "E1": "=OFFSET(INDEX(Sheet1!C1:C2,ROW()-ROW($B$1)+1,1),0,0)",
            }
        },
    )
    assert "Sheet1!B1" not in _graph_keys(workbook, "Sheet1!E1")
    assert "Sheet1!C1" in _graph_keys(workbook, "Sheet1!E1")
    document = bindings_document(
        series_entry(
            "countries",
            "Sheet1!C1:C2",
            layout="series",
            direction="constant",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry("result", "Sheet1!E1", direction="output"),
    )
    package = load_package(generate_inverted(workbook, document), tmp_path)
    assert invoke_public_compute(package, package.compute_result, {}) == pytest.approx(10.0)


def test_same_address_in_value_position_still_requires_a_series(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "value_position.xlsx",
        {"Sheet1": {"B1": 99, "A1": "=ROW($B$1)+$B$1"}},
    )
    document = bindings_document(series_entry("result", "Sheet1!A1", direction="output"))
    with pytest.raises(InvertedTreeExportError, match="Sheet1!B1"):
        generate_inverted(workbook, document)


def test_row_of_index_still_requires_nested_value_refs(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "nested_index.xlsx",
        {"Sheet1": {"C1": 5, "A1": "=ROW(INDEX(Sheet1!C1,1))"}},
    )
    document = bindings_document(series_entry("result", "Sheet1!A1", direction="output"))
    with pytest.raises(InvertedTreeExportError, match="Sheet1!C1"):
        generate_inverted(workbook, document)
