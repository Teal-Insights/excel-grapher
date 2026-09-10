r"""Compare every public output of a generated named package with `FormulaEvaluator`.

Each `compute_*` function runs on the workbook defaults emitted in `data.py`;
every published cell of its result is compared with the evaluator's value for
the authored cell. Numbers match within `atol=1e-6`, `rtol=1e-12`; text and
booleans match exactly; an Excel error matches only the same error code, and a
matched error counts as expected only inside the chart error contract.

    uv run python scripts/named_differential.py PACKAGE_DIR --graph GRAPH.pkl.gz \
        [--blank-ranges sandbox/lic-dsf/lic_dsf_blank_ranges.py] \
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


def _published_cells(function: Any, value: Any) -> dict[tuple[object, ...], object]:
    """Values of the graph-required coordinates, keyed like `__cells__`."""
    if function.__domain__ is None:
        return {next(iter(function.__cells__)): value}
    return {coord: value[coord] for coord in function.__domain__ if coord in function.__cells__}


def run(
    package_dir: Path,
    graph_path: Path,
    contract: Mapping[str, Any] | None,
    blank_ranges: Path | None = None,
) -> dict[str, Any]:
    from excel_grapher.evaluator import FormulaEvaluator
    from excel_grapher.grapher import load_graph
    from excel_grapher.grapher.blank_ranges import load_blank_ranges_module

    sys.path.insert(0, str(package_dir.parent))
    package = importlib.import_module(package_dir.name)
    data = package.data
    started = time.perf_counter()
    graph = load_graph(graph_path)
    blanks = None if blank_ranges is None else load_blank_ranges_module(blank_ranges)
    evaluator = FormulaEvaluator(graph, blank_ranges=blanks)
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
        results[name] = _published_cells(function, value)
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


def run_internals(
    package_dir: Path, graph_path: Path, blank_ranges: Path | None = None
) -> dict[str, Any]:
    """Compare every named formula series (not only outputs) with the evaluator.

    Reports the series whose own value disagrees while every series it reads
    agrees: the root causes of any downstream mismatch.
    """
    from excel_grapher.evaluator import FormulaEvaluator
    from excel_grapher.grapher import load_graph
    from excel_grapher.grapher.blank_ranges import load_blank_ranges_module

    sys.path.insert(0, str(package_dir.parent))
    package = importlib.import_module(package_dir.name)
    data, internals, api = package.data, package.internals, package.api
    graph = load_graph(graph_path)
    blanks = None if blank_ranges is None else load_blank_ranges_module(blank_ranges)
    evaluator = FormulaEvaluator(graph, blank_ranges=blanks)
    inputs = {
        name: getattr(data, name.upper() + "_DEFAULT")
        for name in api.Model.__annotations__
        if hasattr(data, name.upper() + "_DEFAULT")
    }
    model = api.Model(**inputs)
    series = [
        name
        for name in dir(internals)
        if hasattr(getattr(internals, name), "__cells__") and not name.startswith("scan_")
    ]
    values: dict[str, dict[tuple[object, ...], object]] = {}
    failures: dict[str, str] = {}
    input_cells: dict[str, Any] = {}
    for name, default in inputs.items():
        cells = getattr(data, name.upper() + "_CELLS", None)
        if cells is None:
            continue
        input_cells[name] = cells
        if hasattr(default, "domain"):
            values[name] = {
                coord: default[coord]
                for coord in default.domain
                if coord in cells and graph.get_node(cells[coord]) is not None
            }
        else:
            values[name] = {next(iter(cells)): default}
    for name in series:
        try:
            values[name] = _published_cells(getattr(internals, name), getattr(model, name))
        except Exception as exc:  # noqa: BLE001 - reported, not raised
            failures[name] = f"{type(exc).__name__}: {exc}"[:300]

    def provenance_of(name: str) -> Any:
        return input_cells[name] if name in input_cells else getattr(internals, name).__cells__

    addresses = sorted(
        {provenance_of(name)[coord] for name, cells in values.items() for coord in cells}
    )
    expected = evaluator.evaluate(addresses)
    status: dict[str, list[dict[str, Any]]] = {}
    for name, cells in values.items():
        provenance = provenance_of(name)
        bad = []
        for coord, actual in cells.items():
            address = provenance[coord]
            if classify(actual, expected[address]) == "mismatch":
                bad.append(
                    {
                        "coordinate": list(coord),
                        "address": address,
                        "actual": actual,
                        "expected": expected[address],
                    }
                )
        status[name] = bad

    def parameters(name: str) -> tuple[str, ...]:
        if name in input_cells:
            return ()
        code = getattr(internals, name).__code__
        return code.co_varnames[: code.co_kwonlyargcount]

    root_causes = {
        name: bad[:5]
        for name, bad in status.items()
        if bad and all(not status.get(dep) for dep in parameters(name) if dep in status)
    }
    return {
        "series": len(series),
        "failed_series": failures,
        "mismatched_series": sum(1 for bad in status.values() if bad),
        "root_causes": root_causes,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    parser.add_argument("package_dir", type=Path)
    parser.add_argument("--graph", type=Path, required=True)
    parser.add_argument("--contract", type=Path)
    parser.add_argument("--blank-ranges", type=Path, help="Module declaring BLANK_RANGES")
    parser.add_argument("--report", type=Path)
    parser.add_argument(
        "--internals", action="store_true", help="Locate the first divergent named series"
    )
    args = parser.parse_args(argv)
    warnings.simplefilter("ignore")
    if args.internals:
        found = run_internals(args.package_dir.resolve(), args.graph, args.blank_ranges)
        if args.report is not None:
            args.report.write_text(json.dumps(found, indent=2, default=str), encoding="utf-8")
        print(json.dumps({k: found[k] for k in ("series", "mismatched_series")}))
        for name, samples in list(found["root_causes"].items())[:25]:
            print("root cause", name, samples[:2])
        if found["failed_series"]:
            print("failed series", found["failed_series"])
        return 0 if not found["root_causes"] else 1
    contract = json.loads(args.contract.read_text()) if args.contract else None
    report = run(args.package_dir.resolve(), args.graph, contract, args.blank_ranges)
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
