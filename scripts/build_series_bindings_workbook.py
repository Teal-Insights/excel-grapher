"""Generate examples/micro_workbooks/series_bindings.xlsx for binding demos."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "examples" / "micro_workbooks" / "series_bindings.xlsx"


def _write_borvelia_block(ws) -> None:
    ws.write("D1", "Year")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write_number(0, col, year)
    ws.write("A2", "Borvelia")
    ws.write("A3", "Real GDP growth (% per annum)")
    ws.write("A4", "Real interest rate (% per annum)")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, value in enumerate([2.1, 2.2, 2.3, 2.4, 2.5], start=5):
        ws.write_number(2, col, value)
    for col, value in enumerate([1.0, 1.1, 1.2, 1.3, 1.4], start=5):
        ws.write_number(3, col, value)
    for col, value in enumerate([-1.0, -0.5, 0.0, 0.5, 1.0], start=5):
        ws.write_number(4, col, value)


def _write_calendar_year_block(ws) -> None:
    ws.write_number("C10", 1999)
    ws.write_number("D10", 2000)
    ws.write("A11", "Revenue")
    ws.write_number("C11", 100)
    ws.write_number("D11", 200)


def _write_offset_header_block(ws) -> None:
    for col, year in enumerate([1, 2, 3], start=3):
        ws.write_number(14, col, year)
    ws.write("A16", "Revenue")
    for col, value in enumerate([10.0, 20.0, 30.0], start=3):
        ws.write_number(15, col, value)


def _write_scalar_block(ws) -> None:
    ws.write("A25", "Threshold p-value")
    ws.write_number("B25", 0.05)


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(OUT)
    ws = wb.add_worksheet("Sheet1")
    _write_borvelia_block(ws)
    _write_calendar_year_block(ws)
    _write_offset_header_block(ws)
    _write_scalar_block(ws)
    wb.close()
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
