"""Tiny workbook builders for gap reproduction tests."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter


def write_workbook(path: Path, populate) -> Path:
    """Build a workbook via ``populate(workbook)``."""
    workbook = xlsxwriter.Workbook(path)
    populate(workbook)
    workbook.close()
    return path


def write_numbervalue_index_match(path: Path) -> Path:
    """``NUMBERVALUE(TEXT(INDEX(..., MATCH(...))))`` lookup (advanced_formula ``K16`` / #264)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        ws.write_string(4, 10, "PRD-001")
        for row, sku in enumerate(["PRD-001", "PRD-002"], start=5):
            ws.write_string(row - 1, 0, sku)
            ws.write_number(row - 1, 4, 1499.0)
        ws.write_formula(
            15,
            10,
            '=IFERROR(NUMBERVALUE(TEXT(INDEX($E$5:$E$19,MATCH($K$5,$A$5:$A$19,0)),"0.00"),".",","),"N/A")',
            None,
            1499,
        )

    return write_workbook(path, populate)
