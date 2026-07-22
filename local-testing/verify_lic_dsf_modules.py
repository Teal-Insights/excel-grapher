#!/usr/bin/env python3
"""Verify local LIC-DSF codegen packages against Excel cached values.

Compares:

1. ``lic_dsf_2025_08_12`` (cell-only) vs workbook ``data_only`` cache
2. ``lic_dsf_2025_08_12_formula_groups`` vs workbook cache
3. cell-only vs formula-groups (where both produce a value)

Uses Chart Data export targets from ``collect_chart_data_cell_keys``.

Run from the repo root (after codegen)::

    uv run python local-testing/codegen_lic_dsf.py
    uv run python local-testing/codegen_lic_dsf.py --formula-groups
    uv run python local-testing/verify_lic_dsf_modules.py

Exit code 0 only when every comparable target matches across the channels
that successfully produced a value, and both packages compute without error.
"""

from __future__ import annotations

import argparse
import importlib
import sys
import time
from collections.abc import Mapping
from dataclasses import dataclass
from math import isfinite
from pathlib import Path
from typing import Any, cast

import fastpyxl

REPO_ROOT = Path(__file__).resolve().parents[1]
LOCAL_TESTING = Path(__file__).resolve().parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
if str(LOCAL_TESTING) not in sys.path:
    sys.path.insert(0, str(LOCAL_TESTING))

from excel_grapher.core.address_keys import (  # noqa: E402
    normalize_key as normalize_address,
)
from excel_grapher.core.address_keys import parse_address  # noqa: E402
from excel_grapher.evaluator.types import XlError  # noqa: E402
from tests.integration.evaluator.utils.lic_dsf_chart_targets import (  # noqa: E402
    WORKBOOK_PATH,
    cells_in_range,
    collect_chart_data_cell_keys,
    parse_range_spec,
)
from tests.utils.excel_workbook_parity import compare_cached_to_evaluator  # noqa: E402

PACKAGE_CELL = "lic_dsf_2025_08_12"
PACKAGE_GROUPS = "lic_dsf_2025_08_12_formula_groups"
RTOL = 1e-5
ATOL = 1e-9


@dataclass(frozen=True, slots=True)
class Mismatch:
    address: str
    channel: str
    excel: object
    got: object
    detail: str


def _is_finite_number(x: object) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool) and isfinite(float(x))


def _normalize_excel_value(raw: object) -> object:
    if isinstance(raw, str):
        err = XlError.from_text(raw)
        if err is not None:
            return err
    return raw


def _read_workbook_cache(workbook: Path, addresses: list[str]) -> dict[str, object]:
    wb = fastpyxl.load_workbook(
        str(workbook),
        data_only=True,
        read_only=True,
        keep_vba=True,
    )
    try:
        out: dict[str, object] = {}
        for addr in addresses:
            sheet, cell = parse_address(addr)
            if sheet not in wb.sheetnames:
                continue
            val = wb[sheet][cell.replace("$", "")].value
            if val is None:
                continue
            out[normalize_address(addr)] = _normalize_excel_value(val)
        return out
    finally:
        wb.close()


def _expand_target_result(target_key: str, value: object) -> dict[str, object]:
    """Expand a ``compute_all`` entry (cell or row-range list) to cell → value."""
    key = normalize_address(target_key)
    if ":" not in key.split("!", 1)[-1]:
        return {key: value}

    sheet, range_a1 = parse_range_spec(target_key if "!" in target_key else key)
    cells = [normalize_address(c) for c in cells_in_range(sheet, range_a1)]
    if not isinstance(value, list):
        # Unexpected scalar for a range target — attach to first cell only.
        return {cells[0]: value} if cells else {}

    flat: list[object] = []
    for row in value:
        if isinstance(row, list):
            flat.extend(row)
        else:
            flat.append(row)
    out: dict[str, object] = {}
    for cell, item in zip(cells, flat, strict=False):
        out[cell] = item
    return out


def _flatten_compute_all(results: Mapping[str, object]) -> dict[str, object]:
    out: dict[str, object] = {}
    for target, value in results.items():
        out.update(_expand_target_result(target, value))
    return out


def _import_package(name: str) -> Any:
    path = LOCAL_TESTING / name
    if not (path / "__init__.py").is_file():
        raise SystemExit(
            f"Missing package {path}. Generate it with:\n"
            f"  uv run python local-testing/codegen_lic_dsf.py"
            + (" --formula-groups" if name.endswith("formula_groups") else "")
        )
    return importlib.import_module(name)


def _run_compute_all(package: Any, label: str) -> dict[str, object]:
    print(f"\ncompute_all({label})...")
    t0 = time.perf_counter()
    raw = cast(dict[str, object], package.compute_all())
    flat = _flatten_compute_all(raw)
    print(
        f"  done in {time.perf_counter() - t0:.1f}s  "
        f"targets={len(raw)}  cells={len(flat)}"
    )
    return flat


def _normalize_module_value(value: object) -> object:
    """Map embedded-runtime XlError enums onto excel_grapher.XlError by code text."""
    if isinstance(value, XlError):
        return value
    # Generated packages embed their own XlError StrEnum; compare by code string.
    code = getattr(value, "value", None)
    if isinstance(code, str):
        err = XlError.from_text(code)
        if err is not None:
            return err
    if isinstance(value, str):
        err = XlError.from_text(value)
        if err is not None:
            return err
    return value


def _compare_to_excel(
    *,
    channel: str,
    module_values: dict[str, object],
    excel: dict[str, object],
    addresses: list[str],
) -> list[Mismatch]:
    mismatches: list[Mismatch] = []
    for addr in addresses:
        if addr not in excel:
            continue
        if addr not in module_values:
            mismatches.append(
                Mismatch(
                    address=addr,
                    channel=channel,
                    excel=excel[addr],
                    got="<missing>",
                    detail="module did not produce this cell",
                )
            )
            continue
        kind = compare_cached_to_evaluator(
            excel[addr],
            _normalize_module_value(module_values[addr]),
            rtol=RTOL,
            atol=ATOL,
        )
        if kind is not None:
            mismatches.append(
                Mismatch(
                    address=addr,
                    channel=channel,
                    excel=excel[addr],
                    got=module_values[addr],
                    detail=kind.value,
                )
            )
    return mismatches


def _compare_modules(
    cell_only: dict[str, object],
    groups: dict[str, object],
    addresses: list[str],
) -> list[Mismatch]:
    mismatches: list[Mismatch] = []
    for addr in addresses:
        if addr not in cell_only or addr not in groups:
            continue
        a = _normalize_module_value(cell_only[addr])
        b = _normalize_module_value(groups[addr])
        if a == b:
            continue
        if _is_finite_number(a) and _is_finite_number(b):
            scale = max(abs(float(a)), abs(float(b)), 1.0)
            if abs(float(a) - float(b)) <= max(ATOL, RTOL * scale):
                continue
            detail = "numeric_drift"
        else:
            detail = "value_mismatch"
        mismatches.append(
            Mismatch(
                address=addr,
                channel="cell_only↔formula_groups",
                excel=a,
                got=b,
                detail=detail,
            )
        )
    return mismatches


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=WORKBOOK_PATH,
        help="LIC-DSF workbook path (data_only cached values)",
    )
    parser.add_argument(
        "--max-print",
        type=int,
        default=30,
        help="Max mismatches to print per channel",
    )
    args = parser.parse_args()

    workbook = Path(args.workbook)
    if not workbook.is_absolute():
        workbook = (REPO_ROOT / workbook).resolve()
    if not workbook.is_file():
        raise SystemExit(f"Workbook not found: {workbook}")

    targets = collect_chart_data_cell_keys()
    addresses = [normalize_address(t) for t in targets]
    print(f"workbook: {workbook}")
    print(f"targets:  {len(addresses)} Chart Data cells")

    print("\nReading Excel cached values (data_only=True)...")
    t0 = time.perf_counter()
    excel = _read_workbook_cache(workbook, addresses)
    print(
        f"  done in {time.perf_counter() - t0:.1f}s  "
        f"cached comparable values={len(excel)}"
    )

    cell_pkg = _import_package(PACKAGE_CELL)
    groups_pkg = _import_package(PACKAGE_GROUPS)

    cell_error: str | None = None
    groups_error: str | None = None
    cell_values: dict[str, object] = {}
    groups_values: dict[str, object] = {}

    try:
        cell_values = _run_compute_all(cell_pkg, PACKAGE_CELL)
    except Exception as exc:  # noqa: BLE001 — report and continue
        cell_error = f"{type(exc).__name__}: {exc}"
        print(f"  FAILED: {cell_error}")

    try:
        groups_values = _run_compute_all(groups_pkg, PACKAGE_GROUPS)
    except Exception as exc:  # noqa: BLE001 — report and continue
        groups_error = f"{type(exc).__name__}: {exc}"
        print(f"  FAILED: {groups_error}")

    mismatches: list[Mismatch] = []
    if cell_error is None:
        mismatches.extend(
            _compare_to_excel(
                channel="cell_only↔excel",
                module_values=cell_values,
                excel=excel,
                addresses=addresses,
            )
        )
    if groups_error is None:
        mismatches.extend(
            _compare_to_excel(
                channel="formula_groups↔excel",
                module_values=groups_values,
                excel=excel,
                addresses=addresses,
            )
        )
    if cell_error is None and groups_error is None:
        mismatches.extend(_compare_modules(cell_values, groups_values, addresses))

    # Summary
    print("\n" + "=" * 72)
    print("Summary")
    print("=" * 72)
    if cell_error:
        print(f"  {PACKAGE_CELL}: ERROR — {cell_error}")
    else:
        n = sum(1 for m in mismatches if m.channel == "cell_only↔excel")
        print(f"  {PACKAGE_CELL} ↔ excel: {n} mismatches")
    if groups_error:
        print(f"  {PACKAGE_GROUPS}: ERROR — {groups_error}")
    else:
        n = sum(1 for m in mismatches if m.channel == "formula_groups↔excel")
        print(f"  {PACKAGE_GROUPS} ↔ excel: {n} mismatches")
    if cell_error is None and groups_error is None:
        n = sum(1 for m in mismatches if m.channel == "cell_only↔formula_groups")
        print(f"  cell_only ↔ formula_groups: {n} mismatches")

    if mismatches:
        print(f"\nFirst mismatches (up to {args.max_print}):")
        for m in mismatches[: args.max_print]:
            print(
                f"  {m.address}  [{m.channel}]  {m.detail}  "
                f"excel/left={m.excel!r}  got/right={m.got!r}"
            )
        if len(mismatches) > args.max_print:
            print(f"  … {len(mismatches) - args.max_print} more")

    ok = (
        cell_error is None
        and groups_error is None
        and not mismatches
    )
    if ok:
        print("\nAll checks passed.")
        raise SystemExit(0)

    print("\nVerification failed.")
    raise SystemExit(1)


if __name__ == "__main__":
    main()
