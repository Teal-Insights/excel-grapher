r"""Compare every public output of a generated named package with `FormulaEvaluator`.

Each `compute_*` function runs on the workbook defaults emitted in `data.py`;
every published cell of its result is compared with the evaluator's value for
the authored cell. Numbers match within `atol=1e-6`, `rtol=1e-12`; text and
booleans match exactly; an Excel error matches only the same error code, and a
matched error counts as expected only inside the chart error contract.

    uv run python scripts/named_differential.py PACKAGE_DIR --graph GRAPH.pkl.gz \
        [--contract plans/named-axis-chart-error-contract.json] [--report out.json]
"""

from __future__ import annotations

import argparse
import importlib
import json
import math
import sys
import time
import warnings
from collections.abc import Mapping
from pathlib import Path
from typing import Any

ATOL = 1e-6
RTOL = 1e-12
ERROR_CODES = frozenset({"#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NAME?", "#NUM!", "#NULL!"})


def _number(value: object) -> float | None:
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return 0.0
    return None


def classify(actual: object, expected: object) -> str:
    """Return `match`, `matched_error`, or `mismatch` for one cell."""
    if (
        isinstance(actual, str)
        and actual in ERROR_CODES
        or (isinstance(expected, str) and expected in ERROR_CODES)
    ):
        return "matched_error" if actual == expected else "mismatch"
    left, right = _number(actual), _number(expected)
    if left is not None and right is not None:
        return "match" if math.isclose(left, right, rel_tol=RTOL, abs_tol=ATOL) else "mismatch"
    return "match" if actual == expected else "mismatch"


def expected_error(address: str, code: str, contract: Mapping[str, Any] | None) -> bool:
    """True when the chart contract allows a matched `#N/A` at `address`."""
    if contract is None or code != "#N/A":
        return False
    sheet, _, cell = address.rpartition("!")
    return sheet.strip("'") == "Chart Data" and cell in contract.get("formulas", {})


def run(package_dir: Path, graph_path: Path, contract: Mapping[str, Any] | None) -> dict[str, Any]:
    from excel_grapher.evaluator import FormulaEvaluator
    from excel_grapher.grapher import load_graph

    sys.path.insert(0, str(package_dir.parent))
    package = importlib.import_module(package_dir.name)
    data = package.data
    started = time.perf_counter()
    graph = load_graph(graph_path)
    evaluator = FormulaEvaluator(graph)
    timings = {"graph_load_seconds": time.perf_counter() - started}
    started = time.perf_counter()
    results: dict[str, dict[tuple[object, ...], object]] = {}
    failures: dict[str, str] = {}
    for name in sorted(n for n in dir(package) if n.startswith("compute_")):
        function = getattr(package, name)
        kwargs = {
            parameter: getattr(data, parameter.upper() + "_DEFAULT")
            for parameter in function.__code__.co_varnames[: function.__code__.co_kwonlyargcount]
        }
        try:
            value = function(**kwargs)
        except Exception as exc:  # noqa: BLE001 - reported, not raised
            failures[name] = f"{type(exc).__name__}: {exc}"[:300]
            continue
        cells = function.__cells__
        results[name] = (
            {(): value} if function.__domain__ is None else {coord: value[coord] for coord in cells}
        )
    timings["compute_seconds"] = time.perf_counter() - started
    addresses = sorted(
        {
            getattr(package, name).__cells__[coord]
            for name, values in results.items()
            for coord in values
        }
    )
    started = time.perf_counter()
    expected = evaluator.evaluate(addresses)
    timings["evaluate_seconds"] = time.perf_counter() - started
    counts = {"match": 0, "matched_error": 0, "unexpected_matched_error": 0, "mismatch": 0}
    mismatches: list[dict[str, Any]] = []
    unexpected: list[dict[str, Any]] = []
    for name, values in results.items():
        cells = getattr(package, name).__cells__
        for coord, actual in values.items():
            address = cells[coord]
            verdict = classify(actual, expected[address])
            record = {
                "function": name,
                "coordinate": list(coord),
                "address": address,
                "actual": actual,
                "expected": expected[address],
            }
            if verdict == "matched_error" and not expected_error(address, str(actual), contract):
                counts["unexpected_matched_error"] += 1
                unexpected.append(record)
            elif verdict == "mismatch":
                mismatches.append(record)
            counts[verdict] += 1
    return {
        "package_dir": str(package_dir),
        "graph": str(graph_path),
        "tolerances": {"atol": ATOL, "rtol": RTOL},
        "functions": len(results) + len(failures),
        "failed_functions": failures,
        "comparisons": sum(counts.values()),
        "counts": counts,
        "mismatches": mismatches[:200],
        "unexpected_matched_errors": unexpected[:200],
        "timings": timings,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    parser.add_argument("package_dir", type=Path)
    parser.add_argument("--graph", type=Path, required=True)
    parser.add_argument("--contract", type=Path)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args(argv)
    warnings.simplefilter("ignore")
    contract = json.loads(args.contract.read_text()) if args.contract else None
    report = run(args.package_dir.resolve(), args.graph, contract)
    if args.report is not None:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
    print(json.dumps({key: report[key] for key in ("functions", "comparisons", "counts")}))
    if report["failed_functions"]:
        print(f"failed functions: {report['failed_functions']}")
    for record in report["mismatches"][:10]:
        print("mismatch", record)
    for record in report["unexpected_matched_errors"][:10]:
        print("unexpected matched error", record)
    return 0 if not report["mismatches"] and not report["failed_functions"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
