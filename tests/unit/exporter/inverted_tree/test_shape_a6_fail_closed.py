"""Layer A6 — fail closed on unbound formula cells and unverifiable refs."""

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


def test_unbound_formula_in_subgraph_is_export_error(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a6_unbound_formula.xlsx",
        {
            "Inputs": {"A1": 1},
            "Engine": {"A1": "=1+1", "B1": "=Engine!A1+Inputs!A1"},
            "Outputs": {"A1": "=Engine!B1"},
        },
    )
    document = bindings_document(
        series_entry("value", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry("path", "Engine!B1", layout="scalar", direction="internal"),
        series_entry("output_path", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match="Engine!A1"):
        generate_inverted(workbook, document)


def test_bound_series_unbound_leaf_ref_is_export_error(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a6_unbound_leaf.xlsx",
        {
            "Inputs": {"A1": 1, "Z99": 2},
            "Engine": {"A1": "=Inputs!A1+Inputs!Z99"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )
    document = bindings_document(
        series_entry("value", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry("path", "Engine!A1", layout="scalar", direction="internal"),
        series_entry("output_path", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match="not in any bound series"):
        generate_inverted(workbook, document)
