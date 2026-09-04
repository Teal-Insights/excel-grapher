"""Graph-derived access functions for #654's three INDEX/OFFSET shapes (#656)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.access import (
    classify_producer_access,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import (
    inverted_graph_parts,
    oriented_document,
    write_oriented_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a26_index_block import (
    _literal_bindings,
    _literal_sheets,
    _offset_bindings,
    _offset_sheets,
    _slide_bindings,
    _slide_sheets,
)


def _parts(tmp_path: Path, sheets, document, stem: str):
    workbook = write_oriented_workbook(tmp_path / f"{stem}.xlsx", sheets, orientation="horizontal")
    return inverted_graph_parts(workbook, oriented_document(document, "horizontal"))


def test_sliding_index_col_is_static_row_is_dynamic(tmp_path: Path) -> None:
    catalog, _deps, graph = _parts(tmp_path, _slide_sheets(), _slide_bindings(), "af_slide")
    access = classify_producer_access(catalog.get("picked"), catalog.get("block"), catalog, graph)
    assert access.row.kind == "dynamic"
    assert access.col.kind == "static"
    assert access.col.coeff == 1
    assert access.width == catalog.get("block").block_width
    assert access.flat_index_expr("row", "col") == f"row * {access.width} + col"


def test_literal_index_column_is_static_affine(tmp_path: Path) -> None:
    catalog, _deps, graph = _parts(tmp_path, _literal_sheets(), _literal_bindings(), "af_lit")
    access = classify_producer_access(catalog.get("picked"), catalog.get("block"), catalog, graph)
    assert access.row.kind == "dynamic"
    assert access.col.kind == "static"
    assert access.col.coeff == 1
    assert access.col.offset == 0


def test_offset_into_matrix_classifies_axes(tmp_path: Path) -> None:
    catalog, _deps, graph = _parts(tmp_path, _offset_sheets(), _offset_bindings(), "af_off")
    access = classify_producer_access(catalog.get("picked"), catalog.get("block"), catalog, graph)
    assert access.col.kind == "static"
    assert access.col.coeff == 1
    assert access.row.kind in {"static", "dynamic"}


def test_unbound_index_column_still_fails_closed_on_emit(tmp_path: Path) -> None:
    from tests.unit.exporter.inverted_tree.helpers import generate_inverted
    from tests.unit.exporter.inverted_tree.test_shape_a26_index_block import _case_sheets

    sheets = _case_sheets(
        lambda _col, _index: "=INDEX($B$5:$D$7,MATCH($A$1,$A$5:$A$7,0),$Z$9)",
        extra={"Z9": 2},
    )
    workbook = write_oriented_workbook(
        tmp_path / "af_unbound.xlsx", sheets, orientation="horizontal"
    )
    with pytest.raises(InvertedTreeExportError, match="Engine!B2") as exc:
        generate_inverted(workbook, _literal_bindings())
    assert "INDEX column" in str(exc.value) or "block" in str(exc.value)
