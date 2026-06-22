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


def write_software_revenue_sumproduct(path: Path) -> Path:
    """Category-filtered revenue sum (advanced_formula ``K21`` / issue #267)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        categories = (["Software", "Hardware"] * 7) + ["Software"]
        prices = [1499, 800, 1200, 600, 999, 750, 1100, 500, 1300, 700, 1600, 900, 1400, 850, 1500]
        for row, (category, price) in enumerate(zip(categories, prices, strict=True), start=5):
            ws.write_string(row - 1, 2, category)
            ws.write_number(row - 1, 4, price)
        ws.write_formula(
            20,
            10,
            '=SUMPRODUCT(($C$5:$C$19="Software")*$E$5:$E$19)',
            None,
            10598,
        )

    return write_workbook(path, populate)


def write_sumproduct_category_filter(path: Path) -> Path:
    r"""``SUMPRODUCT((range="label")*values)`` (financial_model ``I14``)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        categories = ["Software", "Hardware", "Software", "Hardware"] * 2
        values = [100, 50, 200, 75, 150, 60, 180, 90]
        for row, (category, value) in enumerate(zip(categories, values, strict=True), start=5):
            ws.write_string(row - 1, 2, category)
            ws.write_number(row - 1, 5, value)
        ws.write_formula(
            13,
            8,
            '=SUMPRODUCT((C5:C12="Software")*F5:F12)',
            None,
            630,
        )

    return write_workbook(path, populate)


def write_sumproduct_threshold_count(path: Path) -> Path:
    """``SUMPRODUCT((range>threshold)*1)`` (financial_model ``I18``)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        for row, value in enumerate([100, 250, 200, 75, 300, 60, 180, 400], start=5):
            ws.write_number(row - 1, 3, value)
        ws.write_formula(17, 8, "=SUMPRODUCT((D5:D12>200)*1)", None, 3)

    return write_workbook(path, populate)


def write_sumproduct_price_threshold_k24(path: Path) -> Path:
    """``SUMPRODUCT(($E$5:$E$19>1000)*1)`` (advanced_formula ``K24`` / issue #265)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        prices = [1499, 800, 1200, 600, 999, 750, 1100, 500, 1300, 700, 1600, 900, 1400, 850, 1500]
        for row, price in enumerate(prices, start=5):
            ws.write_number(row - 1, 4, price)
        ws.write_formula(23, 10, "=SUMPRODUCT(($E$5:$E$19>1000)*1)", None, 7)

    return write_workbook(path, populate)


def write_large_numeric_sumproduct(path: Path, *, rows: int = 500) -> Path:
    """``SUMPRODUCT`` over a large numeric range product (Sprint 1 parity fixture)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Data")
        for row in range(rows):
            ws.write_number(row, 0, float((row % 5) + 1))
            ws.write_number(row, 1, 10.0)
        last_row = rows
        ws.write_formula(
            0,
            2,
            f"=SUMPRODUCT(A1:A{last_row}*B1:B{last_row})",
            None,
            15_000.0,
        )

    return write_workbook(path, populate)


def write_large_string_criteria_sumproduct(path: Path, *, rows: int = 2_000) -> Path:
    """``SUMPRODUCT`` over large string-equality criteria (operator parity fixture)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Data")
        for row in range(rows):
            category = "Software" if row % 2 == 0 else "Hardware"
            ws.write_string(row, 0, category)
            ws.write_number(row, 1, 500.0)
        last_row = rows
        expected = (rows // 2) * 500.0
        ws.write_formula(
            0,
            2,
            f'=SUMPRODUCT((A1:A{last_row}="Software")*B1:B{last_row})',
            None,
            expected,
        )

    return write_workbook(path, populate)
