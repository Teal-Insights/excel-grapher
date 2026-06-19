"""Unit tests for ``_xludf.`` workbook/formula rewrite helpers."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter
from fastpyxl import load_workbook

from tests.integration.utils.rewrite_xludf_workbook import (
    rewrite_formula_to_xludf,
    write_xludf_workbook_copy,
)


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        (
            '=IFNA(XLOOKUP(1,A1:A3,B1:B3),"x")',
            '=_xludf.IFNA(_xludf.XLOOKUP(1,A1:A3,B1:B3),"x")',
        ),
        (
            '=_xlfn.NUMBERVALUE("1,234.56", ".", ",")',
            '=_xludf.NUMBERVALUE("1,234.56", ".", ",")',
        ),
        (
            '=_xludf.IFNA(A1,"x")',
            '=_xludf.IFNA(A1,"x")',
        ),
        (
            "=SUM(A1:A3)",
            "=SUM(A1:A3)",
        ),
        ("not a formula", "not a formula"),
    ],
)
def test_rewrite_formula_to_xludf(source: str, expected: str) -> None:
    assert rewrite_formula_to_xludf(source) == expected


def test_rewrite_formula_to_xludf_is_idempotent() -> None:
    formula = '=IFNA(XLOOKUP(1,A1:A3,B1:B3),"x")'
    once = rewrite_formula_to_xludf(formula)
    twice = rewrite_formula_to_xludf(once)
    assert once == twice


def test_write_xludf_workbook_copy_rewrites_formulas(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    destination = tmp_path / "prefix_xludf.xlsx"
    workbook = xlsxwriter.Workbook(source)
    worksheet = workbook.add_worksheet("Lookups")
    worksheet.write_string(0, 1, "hit")
    worksheet.write_formula(0, 0, '=IFNA(B1,"x")')
    workbook.close()

    write_xludf_workbook_copy(source, destination, workbook_name="prefix_xludf.xlsx")

    loaded = load_workbook(destination, data_only=False)
    try:
        formula = loaded["Lookups"]["A1"].value
        formula_text = (
            formula if isinstance(formula, str) else getattr(formula, "text", str(formula))
        )
    finally:
        loaded.close()
    assert "_xludf.IFNA" in formula_text


def test_write_xludf_workbook_copy_updates_binding_workbook_field(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    destination = tmp_path / "prefix_xludf.xlsx"
    bindings_dir = tmp_path / "prefix_xludf.bindings"
    bindings_dir.mkdir()
    shard = bindings_dir / "lookups.bindings.yaml"
    shard.write_text("workbook: source.xlsx\nseries: []\n", encoding="utf-8")

    workbook = xlsxwriter.Workbook(source)
    workbook.add_worksheet("S")
    workbook.close()

    write_xludf_workbook_copy(source, destination, workbook_name="prefix_xludf.xlsx")

    assert "workbook: prefix_xludf.xlsx" in shard.read_text(encoding="utf-8")
