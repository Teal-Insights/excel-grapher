"""Build a micro workbook with one demo region per formula-compression rule.

The companion script ``demo_compression_rules.py`` walks the full
``compress_full`` pipeline order from the compression design doc, applying
implemented rules live and illustrating the rest from this workbook.

Run from the repo root::

    uv run python examples/micro_workbooks/build_compression_rules_workbook.py
"""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

OUT = Path(__file__).with_name("compression_rules.xlsx")

# Excel 1-based rows referenced in formulas below.
RULE1_ROW = 5
RULE2_ROW = 10
RULE3_ROW = 14
RULE4_VALUE_ROW = 17
RULE4_START_ROW = 18
TACO_FIRST = 3
TACO_LAST = 7
RF_TAIL_ROW = 11


def _row(excel_row: int) -> int:
    """Map 1-based Excel row to xlsxwriter's 0-based row."""
    return excel_row - 1


def _write_legend(ws) -> None:
    ws.set_column(0, 0, 6)
    ws.set_column(1, 1, 28)
    ws.set_column(2, 2, 72)
    ws.write(0, 0, "Step")
    ws.write(0, 1, "Rule id")
    ws.write(0, 2, "Description")
    rows = [
        (
            1,
            "pass_through",
            "Rewrite formulas that reference pass-through cells (=B5) to the ultimate target.",
        ),
        (
            2,
            "parallel_if_row",
            "Merge contiguous same-row formulas with a shared template into ParallelFormulaNode.",
        ),
        (
            3,
            "constant_folding",
            "Pre-compute literal-only subexpressions (e.g. =2+3 → 5).",
        ),
        (
            "4a",
            "common_subexpression",
            "Cell CSE fixpoint: hoist shared subtrees to _cse! bindings.",
        ),
        (
            "4b",
            "constant_folding",
            "Post-CSE constant fold on formulas that still contain literals.",
        ),
        (
            "5–9",
            "taco_rr … taco_rr_chain",
            "TACO range-pattern compression (RR, RF, FR, FF, RR-Chain) on the Taco sheet.",
        ),
        (
            10,
            "(remaining)",
            "Add any per-cell formulas not absorbed by artifacts to the compressed map.",
        ),
        (
            "4c",
            "common_subexpression",
            "Artifact CSE across ParallelFormulaNode and TacoPatternNode templates.",
        ),
    ]
    for idx, (step, rule_id, description) in enumerate(rows, start=1):
        ws.write(idx, 0, step)
        ws.write(idx, 1, rule_id)
        ws.write(idx, 2, description)


def _write_ext_sheet(wb) -> None:
    ws = wb.add_worksheet("Ext")
    ws.set_column(3, 3, 10)
    ws.set_column(4, 6, 10)
    ws.write(_row(1), 3, "Shared flag")
    ws.write(_row(1), 4, "D col")
    ws.write(_row(1), 5, "E col")
    ws.write(_row(1), 6, "F col")
    ws.write_string(_row(3), 3, "Yes")
    ws.write_number(_row(87), 3, 10.0)
    ws.write_number(_row(87), 4, 20.0)
    ws.write_number(_row(87), 5, 30.0)


def _write_compress_sheet(wb) -> None:
    ws = wb.add_worksheet("Compress")
    ws.set_column(0, 6, 16)

    ws.write(_row(1), 0, "Formula AST compression demos")

    # --- Rule 1: pass-through ---
    ws.write(_row(3), 0, "Rule 1")
    ws.write(_row(3), 1, "pass_through")
    ws.write(_row(4), 0, "A (transit)")
    ws.write(_row(4), 1, "B (source)")
    ws.write(_row(4), 2, "C = A+10")
    ws.write(_row(4), 3, "D = A*2")
    ws.write_number(_row(RULE1_ROW), 1, 42.0)
    ws.write_formula(_row(RULE1_ROW), 0, f"=B{RULE1_ROW}", None, 42.0)
    ws.write_formula(_row(RULE1_ROW), 2, f"=A{RULE1_ROW}+10", None, 52.0)
    ws.write_formula(_row(RULE1_ROW), 3, f"=A{RULE1_ROW}*2", None, 84.0)

    # --- Rule 2: parallel IF row ---
    ws.write(_row(8), 0, "Rule 2")
    ws.write(_row(8), 1, "parallel_if_row")
    ws.write(_row(9), 3, "D")
    ws.write(_row(9), 4, "E")
    ws.write(_row(9), 5, "F")
    for col_letter, col_idx, expected in (
        ("D", 3, 10.0),
        ("E", 4, 20.0),
        ("F", 5, 30.0),
    ):
        ws.write_formula(
            _row(RULE2_ROW),
            col_idx,
            (f'=IF(Ext!$D$3="No",NA(),Ext!{col_letter}87)'),
            None,
            expected,
        )

    # --- Rule 3: constant folding ---
    ws.write(_row(12), 0, "Rule 3")
    ws.write(_row(12), 1, "constant_folding")
    ws.write(_row(13), 0, "=2+3")
    ws.write(_row(13), 1, "=4*4")
    ws.write(_row(13), 2, '="Hello"&" "&"World"')
    ws.write_formula(_row(RULE3_ROW), 0, "=2+3", None, 5.0)
    ws.write_formula(_row(RULE3_ROW), 1, "=4*4", None, 16.0)
    ws.write_formula(_row(RULE3_ROW), 2, '="Hello"&" "&"World"', None, "Hello World")

    # --- Rule 4: cell CSE ---
    ws.write(_row(16), 0, "Rule 4a")
    ws.write(_row(16), 1, "common_subexpression")
    ws.write(_row(RULE4_VALUE_ROW), 1, "B")
    ws.write(_row(RULE4_VALUE_ROW), 2, "C")
    ws.write_number(_row(RULE4_VALUE_ROW), 1, 1.0)
    ws.write_number(_row(RULE4_VALUE_ROW), 2, 2.0)
    ws.write(_row(RULE4_VALUE_ROW + 1), 0, "=(B+C)*2")
    ws.write(_row(RULE4_VALUE_ROW + 2), 0, "=(B+C)*3")
    ws.write(_row(RULE4_VALUE_ROW + 3), 0, "=(B+C)+10")
    for offset, formula, expected in (
        (0, f"=(B{RULE4_VALUE_ROW}+C{RULE4_VALUE_ROW})*2", 6.0),
        (1, f"=(B{RULE4_VALUE_ROW}+C{RULE4_VALUE_ROW})*3", 9.0),
        (2, f"=(B{RULE4_VALUE_ROW}+C{RULE4_VALUE_ROW})+10", 13.0),
    ):
        ws.write_formula(
            _row(RULE4_START_ROW + offset),
            0,
            formula,
            None,
            expected,
        )


def _write_taco_sheet(wb) -> None:
    ws = wb.add_worksheet("Taco")
    ws.set_column(0, 0, 3)
    ws.set_column(1, 16, 14)
    ws.write(_row(1), 0, "TACO rules 5–9 (column autofill)")

    def row(excel_row: int) -> int:
        return excel_row - 1

    # RR (cols B-D)
    ws.write(row(2), 1, "RR")
    ws.write(row(TACO_FIRST - 1), 1, "B")
    ws.write(row(TACO_FIRST - 1), 2, "C")
    ws.write(row(TACO_FIRST - 1), 3, "D=B*C")
    for er in range(TACO_FIRST, TACO_LAST + 1):
        ws.write_number(row(er), 1, float(er - 2))
        ws.write_number(row(er), 2, float((er - 2) * 10))
        ws.write_formula(row(er), 3, f"=B{er}*C{er}", None, float((er - 2) * (er - 2) * 10))

    # RF (cols E-F)
    ws.write(row(2), 4, "RF")
    ws.write(row(TACO_FIRST - 1), 4, "E")
    ws.write(row(TACO_FIRST - 1), 5, "F=SUM(E:$E$11)")
    for er in range(TACO_FIRST, RF_TAIL_ROW + 1):
        ws.write_number(row(er), 4, 10.0)
    for er in range(TACO_FIRST, TACO_LAST + 1):
        partial = 10.0 * (RF_TAIL_ROW - er + 1)
        ws.write_formula(row(er), 5, f"=SUM(E{er}:$E${RF_TAIL_ROW})", None, partial)

    # FR (cols G-H)
    ws.write(row(2), 6, "FR")
    ws.write(row(TACO_FIRST - 1), 6, "G")
    ws.write(row(TACO_FIRST - 1), 7, "H=SUM($G$3:G)")
    for er in range(TACO_FIRST, TACO_LAST + 1):
        ws.write_number(row(er), 6, float(er - 2))
        ws.write_formula(row(er), 7, f"=SUM($G${TACO_FIRST}:G{er})", None, sum(range(er - 2)))

    # FF (cols J-K, lookup M-N)
    ws.write(row(2), 9, "FF")
    ws.write(row(TACO_FIRST - 1), 9, "J")
    ws.write(row(TACO_FIRST - 1), 10, "K=VLOOKUP")
    ws.write(row(TACO_FIRST - 1), 12, "M")
    ws.write(row(TACO_FIRST - 1), 13, "N")
    keys = ["alpha", "beta", "gamma", "delta", "epsilon"]
    values = [100.0, 200.0, 300.0, 400.0, 500.0]
    for i, (key, value) in enumerate(zip(keys, values, strict=True)):
        er = TACO_FIRST + i
        ws.write_string(row(er), 9, key)
        ws.write_string(row(er), 12, key)
        ws.write_number(row(er), 13, value)
    for er in range(TACO_FIRST, TACO_LAST + 1):
        ws.write_formula(
            row(er),
            10,
            f"=VLOOKUP(J{er},$M${TACO_FIRST}:$N${TACO_FIRST + len(keys) - 1},2,FALSE)",
            None,
            values[er - TACO_FIRST],
        )

    # RR-Chain (col P)
    ws.write(row(2), 15, "RR-Chain")
    ws.write(row(TACO_FIRST - 1), 15, "P")
    ws.write_number(row(TACO_FIRST), 15, 1.0)
    for er in range(TACO_FIRST + 1, TACO_LAST + 1):
        ws.write_formula(row(er), 15, f"=P{er - 1}+1", None, float(er - 1))


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(OUT)
    legend = wb.add_worksheet("Legend")
    _write_legend(legend)
    _write_ext_sheet(wb)
    _write_compress_sheet(wb)
    _write_taco_sheet(wb)
    wb.close()
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
