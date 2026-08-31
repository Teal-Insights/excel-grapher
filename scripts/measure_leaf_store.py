#!/usr/bin/env python3
"""Measure exported leaf payload: NodeKey dict vs nested coordinate store (#579).

Reports before/after bytes, occupied leaves vs distinct sheets, generated
module import time, and `xl_range` scan time over a leaf rectangle.

Usage:
    uv run python scripts/measure_leaf_store.py
    uv run python scripts/measure_leaf_store.py --synthetic-leaves 20000
    uv run python scripts/measure_leaf_store.py --json
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, cast

from excel_grapher.core.address_keys import format_cell_key, parse_cell_coords
from excel_grapher.runtime.leaves import LeafStore

_DESCRIPTION = "Measure exported leaf payload: NodeKey dict vs coordinate store (#579)."

_REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = _REPO_ROOT / "examples" / "micro_workbooks" / "taco_patterns.xlsx"
DEFAULT_TARGETS = (
    "Patterns!D3:D7",
    "Patterns!F3:F7",
    "Patterns!H3:H7",
    "Patterns!K3:K3",
    "Patterns!P3:P7",
)
DEFAULT_SYNTHETIC_LEAVES = 5000


def _py_literal(value: object) -> str:
    if value is None:
        return "0"
    if isinstance(value, (bool, int, float, str)):
        return repr(value)
    return "0"


def emit_nodekey_dict_literal(leaves: Sequence[tuple[str, int, int, object]]) -> str:
    """Emit a NodeKey-keyed `DEFAULT_INPUTS` dict (pre-#579 layout)."""
    import fastpyxl.utils.cell

    lines = ["DEFAULT_INPUTS = {"]
    for sheet, row, col, value in leaves:
        col_letter = fastpyxl.utils.cell.get_column_letter(col)
        key = format_cell_key(sheet, col_letter, row)
        lines.append(f"    {key!r}: {_py_literal(value)},")
    lines.append("}")
    return "\n".join(lines) + "\n"


def emit_coordinate_store_literal(leaves: Sequence[tuple[str, int, int, object]]) -> str:
    """Emit a nested `sheet -> {(row, col): value}` `DEFAULT_INPUTS` dict."""
    by_sheet: dict[str, list[tuple[int, int, object]]] = {}
    for sheet, row, col, value in leaves:
        by_sheet.setdefault(sheet, []).append((row, col, value))
    if not by_sheet:
        return "DEFAULT_INPUTS = {}\n"
    lines = ["DEFAULT_INPUTS = {"]
    for sheet in sorted(by_sheet):
        lines.append(f"    {sheet!r}: {{")
        for row, col, value in sorted(by_sheet[sheet], key=lambda item: (item[0], item[1])):
            lines.append(f"        ({row}, {col}): {_py_literal(value)},")
        lines.append("    },")
    lines.append("}")
    return "\n".join(lines) + "\n"


def synthetic_leaves(
    count: int, *, sheet: str = "Sheet1", columns: int = 10
) -> list[tuple[str, int, int, object]]:
    """Build a sparse mixed-type leaf rectangle of `count` occupied cells."""
    leaves: list[tuple[str, int, int, object]] = []
    for i in range(count):
        row = i // columns + 1
        col = i % columns + 1
        kind = i % 4
        if kind == 0:
            value: object = float(i)
        elif kind == 1:
            value = i
        elif kind == 2:
            value = f"s{i}"
        else:
            value = bool(i % 2)
        leaves.append((sheet, row, col, value))
    return leaves


def leaves_from_nodekey_dict(
    values: Mapping[str, object],
) -> list[tuple[str, int, int, object]]:
    out: list[tuple[str, int, int, object]] = []
    for address, value in values.items():
        sheet, row, col = parse_cell_coords(address)
        out.append((sheet, row, col, value))
    return out


def _time_exec(source: str, *, repeats: int) -> float:
    best = float("inf")
    for _ in range(repeats):
        ns: dict[str, Any] = {}
        t0 = time.perf_counter()
        exec(source, ns)
        best = min(best, time.perf_counter() - t0)
        assert "DEFAULT_INPUTS" in ns
    return best


def _time_xl_range_scan(
    store: LeafStore,
    *,
    sheet: str,
    rows: int,
    cols: int,
    repeats: int,
) -> float:
    from excel_grapher.runtime.leaves import MISSING, leaf

    def _scan() -> int:
        hits = 0
        for row in range(1, rows + 1):
            for col in range(1, cols + 1):
                if leaf(store, sheet, row, col) is not MISSING:
                    hits += 1
        return hits

    _scan()
    best = float("inf")
    for _ in range(repeats):
        t0 = time.perf_counter()
        _scan()
        best = min(best, time.perf_counter() - t0)
    return best


def _time_nodekey_range_scan(
    values: Mapping[str, object],
    *,
    sheet: str,
    rows: int,
    cols: int,
    repeats: int,
) -> float:
    import fastpyxl.utils.cell

    from excel_grapher.core.address_keys import format_cell_key

    def _scan() -> int:
        hits = 0
        for row in range(1, rows + 1):
            for col in range(1, cols + 1):
                key = format_cell_key(sheet, fastpyxl.utils.cell.get_column_letter(col), row)
                if key in values:
                    hits += 1
        return hits

    _scan()
    best = float("inf")
    for _ in range(repeats):
        t0 = time.perf_counter()
        _scan()
        best = min(best, time.perf_counter() - t0)
    return best


def measure_leaf_payload(
    leaves: Sequence[tuple[str, int, int, object]],
    *,
    import_repeats: int = 5,
    scan_repeats: int = 5,
) -> dict[str, Any]:
    """Compare NodeKey vs coordinate-store payload metrics."""
    nodekey_src = emit_nodekey_dict_literal(leaves)
    coord_src = emit_coordinate_store_literal(leaves)
    sheets = {sheet for sheet, _row, _col, _value in leaves}
    occupied = len(leaves)

    nodekey_ns: dict[str, Any] = {}
    exec(nodekey_src, nodekey_ns)
    coord_ns: dict[str, Any] = {}
    exec(coord_src, coord_ns)
    nodekey_dict = nodekey_ns["DEFAULT_INPUTS"]
    coord_store = cast(LeafStore, coord_ns["DEFAULT_INPUTS"])

    max_row = max((row for _s, row, _c, _v in leaves), default=0)
    max_col = max((col for _s, _r, col, _v in leaves), default=0)
    sheet = next(iter(sheets)) if len(sheets) == 1 else None

    payload: dict[str, Any] = {
        "occupied_leaves": occupied,
        "distinct_sheets": len(sheets),
        "nodekey_bytes": len(nodekey_src.encode("utf-8")),
        "coordinate_bytes": len(coord_src.encode("utf-8")),
        "bytes_ratio": (
            len(coord_src.encode("utf-8")) / len(nodekey_src.encode("utf-8"))
            if nodekey_src
            else 0.0
        ),
        "nodekey_import_s": _time_exec(nodekey_src, repeats=import_repeats),
        "coordinate_import_s": _time_exec(coord_src, repeats=import_repeats),
        "make_context_overlay_ok": True,
    }
    if sheet is not None and max_row > 0 and max_col > 0:
        payload["xl_range_nodekey_s"] = _time_nodekey_range_scan(
            nodekey_dict, sheet=sheet, rows=max_row, cols=max_col, repeats=scan_repeats
        )
        payload["xl_range_coordinate_s"] = _time_xl_range_scan(
            coord_store, sheet=sheet, rows=max_row, cols=max_col, repeats=scan_repeats
        )
    try:
        from excel_grapher.runtime.leaves import overlay_leaf_inputs

        overlay_leaf_inputs(coord_store, {f"{next(iter(sheets))}!A1": 123})
    except (ValueError, StopIteration):
        payload["make_context_overlay_ok"] = False
    return payload


def _leaves_from_workbook(
    workbook: Path, targets: Sequence[str]
) -> list[tuple[str, int, int, object]]:
    from excel_grapher import create_dependency_graph
    from excel_grapher.exporter.codegen import CodeGenerator

    graph = create_dependency_graph(workbook, list(targets), load_values=True)
    code = CodeGenerator(graph).generate(list(targets))
    ns: dict[str, Any] = {}
    exec(code, ns)
    store = ns["DEFAULT_INPUTS"]
    if store and isinstance(next(iter(store.values())), dict):
        out: list[tuple[str, int, int, object]] = []
        for sheet, cells in store.items():
            for (row, col), value in cells.items():
                out.append((sheet, row, col, value))
        return out
    return leaves_from_nodekey_dict(store)


def _render_table(payload: Mapping[str, Any]) -> str:
    occupied = int(payload["occupied_leaves"])
    sheets = int(payload["distinct_sheets"])
    node_b = int(payload["nodekey_bytes"])
    coord_b = int(payload["coordinate_bytes"])
    lines = [
        "exported leaf payload (#579):",
        f"  occupied leaves:     {occupied:,}",
        f"  distinct sheets:      {sheets:,}",
        f"  NodeKey dict bytes:  {node_b:,}",
        f"  coordinate bytes:    {coord_b:,}",
        f"  bytes ratio:         {float(payload['bytes_ratio']):.3f}",
        f"  NodeKey import:      {float(payload['nodekey_import_s']) * 1000:.3f} ms",
        f"  coordinate import:   {float(payload['coordinate_import_s']) * 1000:.3f} ms",
    ]
    if "xl_range_nodekey_s" in payload:
        lines += [
            f"  NodeKey xl_range:    {float(payload['xl_range_nodekey_s']) * 1000:.3f} ms",
            f"  coordinate xl_range: {float(payload['xl_range_coordinate_s']) * 1000:.3f} ms",
        ]
    lines.append(
        "  NodeKey overlay:     " + ("ok" if payload["make_context_overlay_ok"] else "FAILED")
    )
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    """Print before/after leaf-payload metrics for a workbook or synthetic grid."""
    parser = argparse.ArgumentParser(description=_DESCRIPTION)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help=f"Workbook to export (default: {DEFAULT_WORKBOOK.name})",
    )
    parser.add_argument(
        "--targets",
        nargs="+",
        default=list(DEFAULT_TARGETS),
        help="Sheet-qualified target cells or ranges",
    )
    parser.add_argument(
        "--synthetic-leaves",
        type=int,
        default=0,
        help=(
            "If >0, measure a synthetic mixed-type leaf rectangle instead of "
            f"a workbook (default 0; {DEFAULT_SYNTHETIC_LEAVES} is a useful size)"
        ),
    )
    parser.add_argument("--json", action="store_true", help="Print JSON instead of a table")
    parser.add_argument(
        "--import-repeats",
        type=int,
        default=5,
        help="Best-of-N repeats for import timings (default: 5)",
    )
    args = parser.parse_args(argv)

    if args.synthetic_leaves > 0:
        leaves = synthetic_leaves(args.synthetic_leaves)
        source = f"synthetic:{args.synthetic_leaves}"
    else:
        if not args.workbook.is_file():
            parser.error(f"workbook not found: {args.workbook}")
        leaves = _leaves_from_workbook(args.workbook, args.targets)
        source = str(args.workbook)

    payload = measure_leaf_payload(leaves, import_repeats=args.import_repeats)
    payload["source"] = source
    payload["targets"] = (
        f"synthetic-leaves:{args.synthetic_leaves}"
        if args.synthetic_leaves > 0
        else list(args.targets)
    )
    if args.json:
        json.dump(payload, sys.stdout, indent=2)
        sys.stdout.write("\n")
    else:
        print(_render_table(payload))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
