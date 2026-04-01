#!/usr/bin/env python3
"""Recalculate a workbook and verify Chart Data sensitivity to an input shock.

Relates to `https://github.com/Teal-Insights/excel-grapher/issues/79` (GDP / macro
shock and LO vs Excel recalc). This script formalizes a check that a small change
on a configured input row (default: **Input 3** row **12**, columns **X:AR**)
moves **Chart Data** row **63** across **D:X** by roughly the expected magnitude
(default band: ~5% vs baseline), after full recalculation.

Default shock on each input cell is **multiplicative**: ``new_value = cell_value * (1 + sign * bps)``,
where ``bps`` is ``--bps / 10000`` (e.g. ``--bps 10`` → factor ``1.001`` / ``0.999``).

Backends:
  * ``auto`` — same selection as ``tests.utils.modify_and_recalculate`` (xlwings on
    Windows/macOS, PowerShell/COM on WSL, LibreOffice on Linux when available).
  * ``libreoffice`` / ``excel`` — force one engine (``excel`` is xlwings/COM as above).

Examples::

    uv run python example/report_chart_shock_sensitivity.py \\
        --workbook example/data/lic-dsf-template-2025-08-12.xlsm

    uv run python example/report_chart_shock_sensitivity.py \\
        --workbook example/data/dsf-uga.xlsm \\
        --input-sheet \"Input 6(optional)-Standard Test\" --input-row 12 \\
        --input-start-col X --input-end-col AR --bps-mode absolute --bps 10

Requires a local ``.xlsm`` (not committed in all clones). Install xlwings on
Windows/macOS for the Excel path; LibreOffice 25.8+ for the LO path.
"""

from __future__ import annotations

import argparse
import math
import statistics
import sys
import tempfile
from collections.abc import Iterable, Mapping
from pathlib import Path

import fastpyxl
from fastpyxl.utils.cell import column_index_from_string, get_column_letter

# Repo root on path for ``tests`` imports
_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from tests.utils._helpers import is_wsl  # noqa: E402
from tests.utils.modify_and_recalculate import (  # noqa: E402
    ExcelRecalculationError,
    _modify_and_recalculate_with_libreoffice,
    _modify_and_recalculate_with_powershell,
    _modify_and_recalculate_with_xlwings,
    modify_and_recalculate_workbook,
)

DEFAULT_INPUT_SHEET = "Input 3 - Macro-Debt data(DMX)"
DEFAULT_CHART_SHEET = "Chart Data"
DEFAULT_INPUT_ROW = 12
DEFAULT_INPUT_START_COL = "X"
DEFAULT_INPUT_END_COL = "AR"
DEFAULT_OBS_ROW = 63
DEFAULT_OBS_START_COL = "D"
DEFAULT_OBS_END_COL = "X"


def _col_range_inclusive(start_letter: str, end_letter: str) -> list[str]:
    a = column_index_from_string(start_letter)
    b = column_index_from_string(end_letter)
    if a > b:
        a, b = b, a
    return [get_column_letter(i) for i in range(a, b + 1)]


def _read_numeric_row_slice(
    path: Path,
    sheet_name: str,
    row: int,
    col_letters: Iterable[str],
) -> dict[str, float | None]:
    """Load cached values for ``Sheet!{Col}{row}``; skip non-numeric."""
    out: dict[str, float | None] = {}
    wb = fastpyxl.load_workbook(str(path), data_only=True, read_only=True, keep_vba=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheet_name!r}")
        ws = wb[sheet_name]
        for col in col_letters:
            addr = f"{col}{row}"
            val = ws[addr].value
            if val is None or isinstance(val, bool):
                out[addr] = None
            elif isinstance(val, (int, float)) and math.isfinite(float(val)):
                out[addr] = float(val)
            else:
                out[addr] = None
    finally:
        wb.close()
    return out


def _flat_address(sheet: str, col: str, row: int) -> str:
    if " " in sheet or any(c in sheet for c in "-()"):
        q = sheet.replace("'", "''")
        return f"'{q}'!${col}${row}"
    return f"{sheet}!${col}${row}"


def _collect_input_modifications(
    workbook: Path,
    sheet: str,
    row: int,
    col_letters: list[str],
    bps: float,
    bps_mode: str,
    direction: int,
) -> dict[str, float]:
    """direction: +1 or -1."""
    current = _read_numeric_row_slice(workbook, sheet, row, col_letters)
    mods: dict[str, float] = {}
    for col in col_letters:
        addr = f"{col}{row}"
        v = current.get(addr)
        if v is None:
            continue
        bps_frac = bps / 10_000.0
        if bps_mode == "multiplicative":
            new_v = v * (1.0 + direction * bps_frac)
        elif bps_mode == "rate_decimal":
            new_v = v + direction * bps_frac
        elif bps_mode == "absolute":
            new_v = v + direction * bps
        else:
            raise ValueError(f"Unknown bps_mode: {bps_mode}")
        mods[_flat_address(sheet, col, row)] = new_v
    return mods


def relative_pct_changes(
    baseline: Mapping[str, float | None],
    shocked: Mapping[str, float | None],
    *,
    eps: float = 1e-12,
) -> list[float]:
    """Per-cell relative changes (shocked - baseline) / |baseline| for numeric pairs."""
    out: list[float] = []
    for key in baseline:
        if key not in shocked:
            continue
        b, s = baseline[key], shocked[key]
        if b is None or s is None:
            continue
        if abs(b) < eps:
            continue
        out.append((s - b) / abs(b))
    return out


def summarize_relative_changes(pcts: list[float]) -> dict[str, float | int]:
    if not pcts:
        return {"n": 0, "median_abs_pct": float("nan"), "mean_abs_pct": float("nan")}
    abs_p = [abs(x) for x in pcts]
    return {
        "n": len(pcts),
        "median_abs_pct": float(statistics.median(abs_p)),
        "mean_abs_pct": float(sum(abs_p) / len(abs_p)),
        "min_pct": float(min(pcts)),
        "max_pct": float(max(pcts)),
    }


def _run_recalc(
    backend: str,
    input_path: Path,
    output_path: Path,
    cell_modifications: dict[str, float],
) -> None:
    if backend == "auto":
        modify_and_recalculate_workbook(input_path, output_path, cell_modifications)
        return
    if backend == "libreoffice":
        _modify_and_recalculate_with_libreoffice(input_path, output_path, cell_modifications)
        return
    if backend == "excel":
        if sys.platform in ("win32", "darwin"):
            _modify_and_recalculate_with_xlwings(input_path, output_path, cell_modifications)
            return
        if is_wsl():
            _modify_and_recalculate_with_powershell(input_path, output_path, cell_modifications)
            return
        raise RuntimeError(
            "backend=excel requires Windows, macOS, or WSL (xlwings / COM to Excel)."
        )
    raise ValueError(f"Unknown backend: {backend}")


def _sort_cell_addrs(addrs: Iterable[str]) -> list[str]:
    def key(a: str) -> tuple[int, int]:
        letters = "".join(c for c in a if c.isalpha())
        digits_s = "".join(c for c in a if c.isdigit())
        row = int(digits_s) if digits_s else 0
        return (row, column_index_from_string(letters))

    return sorted(addrs, key=key)


def _print_table(
    title: str,
    baseline: Mapping[str, float | None],
    shocked: Mapping[str, float | None],
) -> None:
    print(f"\n{title}")
    print(f"{'Cell':<8} {'Baseline':>14} {'Shocked':>14} {'Rel chg':>12}")
    for addr in _sort_cell_addrs(baseline.keys()):
        b, s = baseline[addr], shocked.get(addr)
        if b is None or s is None:
            rel = ""
        elif abs(b) < 1e-12:
            rel = "n/a"
        else:
            rel = f"{(s - b) / abs(b) * 100:.2f}%"
        print(
            f"{addr:<8} {str(b) if b is not None else '':>14} "
            f"{str(s) if s is not None else '':>14} {rel:>12}"
        )


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--workbook", type=Path, required=True, help="Source .xlsm path")
    p.add_argument("--input-sheet", default=DEFAULT_INPUT_SHEET, help="Sheet holding input row")
    p.add_argument("--input-row", type=int, default=DEFAULT_INPUT_ROW)
    p.add_argument("--input-start-col", default=DEFAULT_INPUT_START_COL)
    p.add_argument("--input-end-col", default=DEFAULT_INPUT_END_COL)
    p.add_argument("--chart-sheet", default=DEFAULT_CHART_SHEET)
    p.add_argument("--chart-row", type=int, default=DEFAULT_OBS_ROW)
    p.add_argument("--chart-start-col", default=DEFAULT_OBS_START_COL)
    p.add_argument("--chart-end-col", default=DEFAULT_OBS_END_COL)
    p.add_argument(
        "--bps",
        type=float,
        default=10.0,
        help="Basis points for multiplicative/rate_decimal modes (10 => 0.001 fraction)",
    )
    p.add_argument(
        "--bps-mode",
        choices=("multiplicative", "rate_decimal", "absolute"),
        default="multiplicative",
        help="multiplicative (default): v * (1 + sign*bps/10000); rate_decimal: v + sign*bps/10000; absolute: v + sign*bps",
    )
    p.add_argument(
        "--backend",
        choices=("auto", "excel", "libreoffice"),
        default="auto",
        help="Recalculation engine",
    )
    p.add_argument(
        "--expect-median-abs-pct-low",
        type=float,
        default=2.0,
        help="Lower bound on median |relative change| in %% (Chart Data row)",
    )
    p.add_argument(
        "--expect-median-abs-pct-high",
        type=float,
        default=12.0,
        help="Upper bound on median |relative change| in %% (Chart Data row)",
    )
    p.add_argument("--keep-temps", action="store_true", help="Print temp paths instead of deleting")
    args = p.parse_args()

    wb = args.workbook.resolve()
    if not wb.is_file():
        print(f"Workbook not found: {wb}", file=sys.stderr)
        return 2

    input_cols = _col_range_inclusive(args.input_start_col, args.input_end_col)
    chart_cols = _col_range_inclusive(args.chart_start_col, args.chart_end_col)

    tmpdir = Path(tempfile.mkdtemp(prefix="chart-shock-"))
    base_out = tmpdir / "baseline.xlsm"
    up_out = tmpdir / "shock_up.xlsm"
    down_out = tmpdir / "shock_down.xlsm"

    try:
        print(f"Workbook: {wb}")
        print(f"Backend: {args.backend}")
        print(
            f"Input shock: {args.input_sheet!r} row {args.input_row} "
            f"cols {args.input_start_col}:{args.input_end_col} "
            f"({args.bps_mode}, bps={args.bps})"
        )
        print(
            f"Observe: {args.chart_sheet!r} row {args.chart_row} "
            f"cols {args.chart_start_col}:{args.chart_end_col}"
        )

        print("\nRecalculating baseline (no input edits)…")
        _run_recalc(args.backend, wb, base_out, {})

        mods_up = _collect_input_modifications(
            wb, args.input_sheet, args.input_row, input_cols, args.bps, args.bps_mode, +1
        )
        mods_down = _collect_input_modifications(
            wb, args.input_sheet, args.input_row, input_cols, args.bps, args.bps_mode, -1
        )
        if not mods_up:
            print(
                "No numeric input cells found to shock (check sheet name / row / columns).",
                file=sys.stderr,
            )
            return 3

        print(f"Applying +shock to {len(mods_up)} cell(s)…")
        _run_recalc(args.backend, wb, up_out, mods_up)
        print(f"Applying -shock to {len(mods_up)} cell(s)…")
        _run_recalc(args.backend, wb, down_out, mods_down)

        base_obs = _read_numeric_row_slice(base_out, args.chart_sheet, args.chart_row, chart_cols)
        up_obs = _read_numeric_row_slice(up_out, args.chart_sheet, args.chart_row, chart_cols)
        down_obs = _read_numeric_row_slice(down_out, args.chart_sheet, args.chart_row, chart_cols)

        up_pcts = relative_pct_changes(base_obs, up_obs)
        down_pcts = relative_pct_changes(base_obs, down_obs)
        sum_up = summarize_relative_changes(up_pcts)
        sum_down = summarize_relative_changes(down_pcts)

        print("\n--- Summary (relative to baseline recalc) ---")
        print(f"+shock: {sum_up}")
        print(f"-shock: {sum_down}")

        _print_table("+shock vs baseline (Chart Data)", base_obs, up_obs)

        med_up = sum_up["median_abs_pct"] * 100.0
        ok = (
            not math.isnan(med_up)
            and args.expect_median_abs_pct_low <= med_up <= args.expect_median_abs_pct_high
        )
        print(
            f"\nMedian |relative change| (+shock): {med_up:.3f}% "
            f"(expected in [{args.expect_median_abs_pct_low}, {args.expect_median_abs_pct_high}]%)"
        )
        if ok:
            print("PASS: median impact within configured band.")
        else:
            print(
                "FAIL: median impact outside band — adjust inputs, bps-mode, or bounds; "
                "or compare backends (Excel vs LibreOffice).",
                file=sys.stderr,
            )

        if args.keep_temps:
            print(f"\nTemp dir: {tmpdir}")
        else:
            for f in (base_out, up_out, down_out):
                f.unlink(missing_ok=True)
            tmpdir.rmdir()

        return 0 if ok else 1
    except (ExcelRecalculationError, RuntimeError, ImportError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        if args.keep_temps:
            print(f"Temp dir: {tmpdir}", file=sys.stderr)
        return 4


if __name__ == "__main__":
    raise SystemExit(main())
