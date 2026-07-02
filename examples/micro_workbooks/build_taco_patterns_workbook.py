"""Build a demo workbook with one example per TACO compression pattern.

Patterns (head / tail reference style when autofilled **down a column**):

- RR: Relative + Relative — e.g. ``=B2*C2`` (both operands shift per row).
- RF: Relative + Fixed — e.g. ``=SUM(E2:$E$10)`` (head moves, tail pinned).
- FR: Fixed + Relative — e.g. ``=SUM($G$2:G2)`` (running total / YTD).
- FF: Fixed + Fixed — e.g. ``=VLOOKUP(J2,$M$2:$N$6,2,FALSE)`` (same table every row).
- RR-Chain: each row references the cell above in the same column.

Row-autofill counterparts (filled **across a row**) live in
``build_taco_row_patterns_workbook.py`` → ``taco_row_patterns.xlsx``.
"""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

OUT = Path(__file__).with_name("taco_patterns.xlsx")
HEADER_ROW = 2
FIRST = 3
LAST = 7
RF_TAIL_ROW = 11


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(OUT)

    legend = wb.add_worksheet("Legend")
    legend.set_column(0, 0, 12)
    legend.set_column(1, 1, 88)
    legend.write(0, 0, "Pattern")
    legend.write(0, 1, "Meaning")
    for i, (pat, desc) in enumerate(
        [
            (
                "RR",
                "Relative-Relative: precedent head and tail both shift with the formula row. "
                "Example: =B2*C2 filled down.",
            ),
            (
                "RF",
                "Relative-Fixed: head shifts, tail pinned. Example: =SUM(E2:$E$10) filled down.",
            ),
            (
                "FR",
                "Fixed-Relative: head pinned, tail grows. Example: =SUM($G$2:G2) running total.",
            ),
            (
                "FF",
                "Fixed-Fixed: same precedent range every row. Example: VLOOKUP into $M$2:$N$6.",
            ),
            (
                "RR-Chain",
                "RR special case: each row depends on the cell immediately above (P3=P2+1, …).",
            ),
            (
                "Row fixtures",
                "See build_taco_row_patterns_workbook.py for the same five patterns filled right.",
            ),
        ],
        start=1,
    ):
        legend.write(i, 0, pat)
        legend.write(i, 1, desc)

    ws = wb.add_worksheet("Patterns")
    ws.set_column(0, 0, 3)
    ws.set_column(1, 16, 14)

    def row(excel_row: int) -> int:
        """Map 1-based Excel row to xlsxwriter's 0-based row."""
        return excel_row - 1

    # --- RR (cols B-D) ---
    ws.write(row(1), 1, "RR")
    ws.write(row(HEADER_ROW), 1, "B (operand)")
    ws.write(row(HEADER_ROW), 2, "C (operand)")
    ws.write(row(HEADER_ROW), 3, "D = B*C")
    for er in range(FIRST, LAST + 1):
        ws.write_number(row(er), 1, float(er - 1))
        ws.write_number(row(er), 2, float((er - 1) * 10))
        ws.write_formula(row(er), 3, f"=B{er}*C{er}", None, float((er - 1) * (er - 1) * 10))

    # --- RF (cols E-F); data E3:E11, formulas F3:F7 ---
    ws.write(row(1), 4, "RF")
    ws.write(row(HEADER_ROW), 4, "E (data)")
    ws.write(row(HEADER_ROW), 5, "F = SUM(Erow:$E$11)")
    for er in range(FIRST, RF_TAIL_ROW + 1):
        ws.write_number(row(er), 4, 10.0)
    for er in range(FIRST, LAST + 1):
        partial = 10.0 * (RF_TAIL_ROW - er + 1)
        ws.write_formula(row(er), 5, f"=SUM(E{er}:$E${RF_TAIL_ROW})", None, partial)

    # --- FR (cols G-H) ---
    ws.write(row(1), 6, "FR")
    ws.write(row(HEADER_ROW), 6, "G (data)")
    ws.write(row(HEADER_ROW), 7, "H = SUM($G$3:Grow)")
    for er in range(FIRST, LAST + 1):
        ws.write_number(row(er), 6, float(er - 1))
        ws.write_formula(row(er), 7, f"=SUM($G${FIRST}:G{er})", None, sum(range(er - 1)))

    # --- FF (cols J-K, lookup M-N) ---
    ws.write(row(1), 9, "FF")
    ws.write(row(HEADER_ROW), 9, "J (key)")
    ws.write(row(HEADER_ROW), 10, "K = VLOOKUP")
    ws.write(row(HEADER_ROW), 12, "M (key)")
    ws.write(row(HEADER_ROW), 13, "N (value)")
    keys = ["alpha", "beta", "gamma", "delta", "epsilon"]
    values = [100.0, 200.0, 300.0, 400.0, 500.0]
    for i, (k, v) in enumerate(zip(keys, values, strict=True)):
        er = FIRST + i
        ws.write_string(row(er), 9, k)
        ws.write_string(row(er), 12, k)
        ws.write_number(row(er), 13, v)
    for er in range(FIRST, LAST + 1):
        ws.write_formula(
            row(er),
            10,
            f"=VLOOKUP(J{er},$M${FIRST}:$N${FIRST + len(keys) - 1},2,FALSE)",
            None,
            values[er - FIRST],
        )

    # --- RR-Chain (col P) ---
    ws.write(row(1), 15, "RR-Chain")
    ws.write(row(HEADER_ROW), 15, "P (chain)")
    ws.write_number(row(FIRST), 15, 1.0)
    for er in range(FIRST + 1, LAST + 1):
        ws.write_formula(row(er), 15, f"=P{er - 1}+1", None, float(er - 1))

    wb.close()
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
