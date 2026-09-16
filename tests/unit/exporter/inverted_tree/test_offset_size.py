"""OFFSET height/width must fail closed at inverted-tree emit."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)


def test_offset_height_width_raises_inverted_tree_export_error(tmp_path: Path) -> None:
    """`OFFSET` size args are not a scalar step; emit must refuse them."""
    workbook = write_workbook(
        tmp_path / "offset_size.xlsx",
        {"S": {"A1": 1.0, "C1": "=SUM(OFFSET(A1,0,0,2,2))"}},
    )
    document = bindings_document(
        series_entry("anchor", "S!A1", layout="scalar", direction="constant"),
        series_entry("out", "S!C1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match="height/width"):
        generate_inverted(workbook, document)
