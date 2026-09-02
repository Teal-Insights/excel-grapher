"""Finding 2 — mixed formula shapes in one bound series fail closed."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)


def test_mixed_member_formulas_fail_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a12_shapes.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=1",
                "B2": "=A2*2",
                "C2": "=B2+100",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )
    with pytest.raises(InvertedTreeExportError, match="formula shape"):
        generate_inverted(workbook, document)
