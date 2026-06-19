"""Tiny workbook builders for gap reproduction tests."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter


def write_workbook(path: Path, populate) -> Path:
    """Build a one-sheet or multi-sheet workbook via ``populate(workbook)``."""
    workbook = xlsxwriter.Workbook(path)
    populate(workbook)
    workbook.close()
    return path


def write_index_match_best_week(path: Path) -> Path:
    """``INDEX(C3:V3,MATCH(MAX(C4:V4),C4:V4,0))`` on a small grid."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Aggregate Stats")
        for offset, value in enumerate([10.0, 30.0, 20.0, 40.0, 15.0]):
            col = 2 + offset
            ws.write_number(2, col, value)
            ws.write_number(3, col, value)
        ws.write_formula(27, 3, "=INDEX(C3:V3,MATCH(MAX(C4:V4),C4:V4,0))")

    return write_workbook(path, populate)


def write_text_index_match(path: Path) -> Path:
    r"""``TEXT(INDEX(...,MATCH(MAX(...))),"0")`` revenue-summary shape."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Revenue Model")
        for col, year in enumerate([2025, 2026, 2027, 2028, 2029], start=1):
            ws.write_number(3, col, year)
            ws.write_number(4, col, 1000 + col * 100)
        ws.write_formula(
            21,
            1,
            '=TEXT(INDEX(B4:F4,MATCH(MAX(B5:F5),B5:F5,0)),"0")',
            None,
            "2029",
        )

    return write_workbook(path, populate)


def write_sumproduct_std_dev(path: Path) -> Path:
    """Sample std-dev via ``SUMPRODUCT`` variance (financial_model ``F6``)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Statistical Analysis")
        for row, value in enumerate(
            [312, 400, 280, 350, 290, 310, 330, 295, 305, 315, 325, 340],
            start=5,
        ):
            ws.write_number(row - 1, 1, value)
        ws.write_formula(
            5,
            5,
            "=IFERROR(SUMPRODUCT((B5:B16-AVERAGE(B5:B16))^2)/COUNT(B5:B16),NA())^0.5",
            None,
            102.9,
        )

    return write_workbook(path, populate)


def write_sumproduct_category_filter(path: Path) -> Path:
    r"""``SUMPRODUCT((range="label")*values)`` (financial_model ``I14``–``I16``)."""

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
            640,
        )
        ws.write_formula(
            14,
            8,
            '=SUMPRODUCT((C5:C12="Hardware")*F5:F12)',
            None,
            275,
        )
        ws.write_formula(
            15,
            8,
            '=SUMPRODUCT((C5:C12="Accessory")*F5:F12)',
            None,
            0,
        )

    return write_workbook(path, populate)


def write_sumproduct_row_weighted_category_count(path: Path) -> Path:
    """``SUMPRODUCT((ROW(...)-ROW(...)+1)*(range="label"))`` (financial_model ``I19``)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        categories = ["Software", "Hardware"] * 4
        for row, category in enumerate(categories, start=5):
            ws.write_string(row - 1, 2, category)
        ws.write_formula(
            18,
            8,
            '=SUMPRODUCT((ROW(A5:A12)-ROW(A5)+1)*(C5:C12="Software"))',
            None,
            12,
        )

    return write_workbook(path, populate)


def write_sumproduct_threshold_count(path: Path) -> Path:
    """``SUMPRODUCT((range>threshold)*1)`` (financial_model ``I18``)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        for row, value in enumerate([100, 50, 200, 75, 150, 60, 180, 90], start=5):
            ws.write_number(row - 1, 3, value)
        ws.write_formula(17, 8, "=SUMPRODUCT((D5:D12>200)*1)", None, 3)

    return write_workbook(path, populate)


def write_software_revenue_sumproduct(path: Path) -> Path:
    """Category-filtered revenue sum (advanced_formula ``K21``)."""

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
            5000,
        )

    return write_workbook(path, populate)


def write_numbervalue_index_match(path: Path) -> Path:
    """``NUMBERVALUE(TEXT(INDEX(...,MATCH(...)),...))`` (advanced_formula ``K16``)."""

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


def write_sumproduct_price_threshold_k24(path: Path) -> Path:
    """``SUMPRODUCT(($E$5:$E$19>1000)*1)`` (advanced_formula ``K24``)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        for row, value in enumerate([1499, 800, 1200, 600, 999], start=5):
            ws.write_number(row - 1, 4, value)
        ws.write_formula(23, 10, "=SUMPRODUCT(($E$5:$E$19>1000)*1)", None, 3)

    return write_workbook(path, populate)


def write_statistical_normdist_panel(path: Path) -> Path:
    """NORMDIST panel whose z-score/CDF depend on sample std-dev (financial_model ``F6``–``H16``)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Statistical Analysis")
        for row, value in enumerate(
            [312, 400, 280, 350, 290, 310, 330, 295, 305, 315, 325, 340],
            start=5,
        ):
            ws.write_number(row - 1, 1, value)
        ws.write_formula(4, 5, "=AVERAGE(B5:B16)", None, 318.75)
        ws.write_formula(
            5,
            5,
            "=IFERROR(SUMPRODUCT((B5:B16-AVERAGE(B5:B16))^2)/COUNT(B5:B16),NA())^0.5",
            None,
            102.9,
        )
        ws.write_number(15, 4, 312)
        ws.write_formula(15, 5, "=IFERROR((E16-$F$5)/$F$6,NA())", None, -0.066)
        ws.write_formula(
            15,
            6,
            "=IFERROR(NORMDIST(E16,$F$5,$F$6,TRUE()),NA())",
            None,
            0.47,
        )
        ws.write_formula(
            15,
            7,
            '=IF(ISNUMBER(F16),IF(F16>0,"Above","Below"),"N/A")',
            None,
            "Below",
        )

    return write_workbook(path, populate)


def write_text_currency_thousands(path: Path) -> Path:
    r"""``TEXT(value,"$#,##0")&"K"`` enterprise-value label (financial_model ``B24``)."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("DCF Valuation")
        for col, value in enumerate([1000, 800, 900, 543], start=1):
            ws.write_number(12, col, value)
        ws.write_formula(17, 1, "=SUM(B13:E13)", None, 3243)
        ws.write_formula(
            23,
            1,
            '=IFERROR(TEXT(B18,"$#,##0")&"K","N/A")',
            None,
            "$3,243K",
        )

    return write_workbook(path, populate)


def write_vlookup_false(path: Path) -> Path:
    """Minimal ``VLOOKUP(...,FALSE())`` table for modular export gap."""

    def populate(workbook: xlsxwriter.Workbook) -> None:
        ws = workbook.add_worksheet("Product Lookup")
        ws.write_string(4, 8, "P003")
        for row, (product_id, name, price) in enumerate(
            [("P001", "Alpha", 10.0), ("P002", "Beta", 20.0), ("P003", "Gamma", 30.0)],
            start=5,
        ):
            ws.write_string(row - 1, 0, product_id)
            ws.write_string(row - 1, 1, name)
            ws.write_number(row - 1, 2, price)
        ws.write_formula(
            4,
            9,
            '=IFERROR(VLOOKUP(I5,A5:C7,2,FALSE()),"Not Found")',
            None,
            "Gamma",
        )

    return write_workbook(path, populate)
