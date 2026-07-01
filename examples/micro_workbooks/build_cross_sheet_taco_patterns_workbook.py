"""Build a two-sheet demo workbook for cross-sheet TACO compression patterns.

``Data`` holds inputs; ``Report`` holds formulas that reference ``Data``.
"""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

OUT = Path(__file__).with_name("cross_sheet_taco_patterns.xlsx")
FIRST = 3
LAST = 7
RF_TAIL_ROW = 11


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(OUT)

    data = wb.add_worksheet("Data")
    report = wb.add_worksheet("Report")

    def row(excel_row: int) -> int:
        return excel_row - 1

    keys = ["alpha", "beta", "gamma", "delta", "epsilon"]
    values = [100.0, 200.0, 300.0, 400.0, 500.0]

    for er in range(FIRST, RF_TAIL_ROW + 1):
        data.write_number(row(er), 4, 10.0)  # E

    for er in range(FIRST, LAST + 1):
        data.write_number(row(er), 1, float(er - 1))  # B
        data.write_number(row(er), 2, float((er - 1) * 10))  # C
        data.write_number(row(er), 6, float(er - 1))  # G

    for i, (k, v) in enumerate(zip(keys, values, strict=True)):
        er = FIRST + i
        data.write_string(row(er), 12, k)  # M
        data.write_number(row(er), 13, v)  # N

    for er in range(FIRST, LAST + 1):
        report.write_formula(row(er), 3, f"=Data!B{er}*Data!C{er}")  # RR -> D
        report.write_formula(row(er), 5, f"=SUM(Data!E{er}:Data!$E${RF_TAIL_ROW})")  # RF -> F
        report.write_formula(row(er), 7, f"=SUM(Data!$G${FIRST}:Data!G{er})")  # FR -> H
        report.write_string(row(er), 9, keys[er - FIRST])  # J keys
        report.write_formula(
            row(er),
            10,
            f"=VLOOKUP(J{er},Data!$M${FIRST}:Data!$N${FIRST + len(keys) - 1},2,FALSE)",
        )

    wb.close()
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
