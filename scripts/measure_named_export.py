#!/usr/bin/env python3
r"""Measure a named-axis standalone export: size, structure, lowering, timing.

The sandbox layout is a directory holding `workbook.xlsm` (or `.xlsx`), a
`bindings/` shard directory, a constraints module, a blank-ranges module, and
`targets.json`. Every input and generated file is hashed so a report can be
reproduced from the recorded revision.

Usage:
    uv run python scripts/measure_named_export.py sandbox/lic-dsf \\
        --graph sandbox/lic-dsf/graph/<cache-key>.pkl.gz \\
        --out-dir build/lic_dsf_named --report build/lic_dsf_named.json
    uv run python scripts/measure_named_export.py sandbox/lic-dsf \\
        --package-dir build/lic_dsf_named --report build/lic_dsf_named.json

Without `--graph` the dependency graph is extracted from the workbook with the
constraints, blank ranges, and targets found in the sandbox. `--package-dir`
skips generation and measures an existing package. `--reconcile-bindings`
drops output bindings that only cover constant target cells and the later of
two formula-direction bindings claiming one cell; every drop is reported.
"""

from __future__ import annotations

import argparse
import ast
import collections
import hashlib
import json
import os
import platform
import statistics
import subprocess
import sys
import time
import warnings
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

if TYPE_CHECKING:
    from excel_grapher.series_bindings.types import WorkbookSeriesBindings

_REPO_ROOT = Path(__file__).resolve().parents[1]
_GENERATED_MODULES = (
    "__init__.py",
    "api.py",
    "validation.py",
    "internals.py",
    "data.py",
    "runtime.py",
    "tensor.py",
    "provenance.py",
    "excel.py",
)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _git_revision() -> dict[str, str]:
    def run(*args: str) -> str:
        try:
            return subprocess.run(
                ["git", *args], cwd=_REPO_ROOT, capture_output=True, text=True, check=True
            ).stdout.strip()
        except (OSError, subprocess.CalledProcessError):
            return "unknown"

    return {
        "head": run("rev-parse", "HEAD"),
        "branch": run("rev-parse", "--abbrev-ref", "HEAD"),
        "dirty": "yes" if run("status", "--porcelain", "--untracked-files=no") else "no",
    }


# ---------------------------------------------------------------------------
# Inputs
# ---------------------------------------------------------------------------


def _sandbox_paths(args: argparse.Namespace) -> dict[str, Path]:
    root = Path(args.sandbox)
    workbook = args.workbook or next(
        (path for path in sorted(root.glob("*.xls[xm]")) if path.is_file()),
        None,
    )
    if workbook is None:
        raise SystemExit(f"no workbook.xlsm or workbook.xlsx in {root}")
    constraints = args.constraints or next(iter(sorted(root.glob("*constraints.py"))), None)
    blank_ranges = args.blank_ranges or next(iter(sorted(root.glob("*blank_ranges.py"))), None)
    paths = {
        "workbook": Path(workbook),
        "bindings": Path(args.bindings or root / "bindings"),
        "targets": Path(args.targets or root / "targets.json"),
    }
    if constraints is not None:
        paths["constraints"] = Path(constraints)
    if blank_ranges is not None:
        paths["blank_ranges"] = Path(blank_ranges)
    if args.graph is not None:
        paths["graph"] = Path(args.graph)
    return paths


def _input_hashes(paths: Mapping[str, Path]) -> dict[str, Any]:
    hashes: dict[str, Any] = {}
    for name, path in paths.items():
        if path.is_dir():
            hashes[name] = {
                child.name: _sha256(child) for child in sorted(path.iterdir()) if child.is_file()
            }
        elif path.is_file():
            hashes[name] = _sha256(path)
    return hashes


def _load_graph(paths: Mapping[str, Path], timings: dict[str, float]) -> Any:
    from excel_grapher.grapher.graph_pickle import load_graph

    if "graph" in paths:
        started = time.perf_counter()
        graph = load_graph(paths["graph"])
        timings["graph_load_seconds"] = time.perf_counter() - started
        return graph
    from excel_grapher.grapher import create_dependency_graph
    from excel_grapher.grapher.blank_ranges import load_blank_ranges_module
    from excel_grapher.grapher.constraints import dynamic_refs_from_path

    sys.path.insert(0, str(paths["workbook"].resolve().parent))
    dynamic_refs = dynamic_refs_from_path(paths["constraints"]) if "constraints" in paths else None
    blank = load_blank_ranges_module(paths["blank_ranges"]) if "blank_ranges" in paths else None
    if paths["targets"].is_file():
        targets = json.loads(paths["targets"].read_text(encoding="utf-8"))
    else:
        from excel_grapher.series_bindings.load import load_series_bindings
        from excel_grapher.series_bindings.workflow import all_series_targets

        targets = all_series_targets(
            load_series_bindings(paths["bindings"]), workbook=paths["workbook"]
        )
    started = time.perf_counter()
    graph = create_dependency_graph(
        paths["workbook"],
        targets,
        load_values=True,
        dynamic_refs=dynamic_refs,
        use_cached_dynamic_refs=dynamic_refs is None,
        capture_dependency_provenance=True,
        blank_ranges=blank,
    )
    timings["graph_extract_seconds"] = time.perf_counter() - started
    return graph


def numeric_text_measure(graph: Any, entry: Mapping[str, Any], cells: Sequence[str]) -> bool:
    """True when a text-typed measure binds only numeric workbook cells."""
    measure = entry.get("structure", {}).get("measure", {})
    if measure.get("dtype") not in {"string", "str"}:
        return False
    values = []
    for cell in cells:
        node = graph.get_node(cell)
        if node is None:
            return False
        values.append(node.value)
    return bool(values) and all(
        isinstance(value, int | float) and not isinstance(value, bool) for value in values
    )


def fractional_int_measure(graph: Any, entry: Mapping[str, Any], cells: Sequence[str]) -> bool:
    """True when an integer-typed measure binds a cell holding a fractional value."""
    measure = entry.get("structure", {}).get("measure", {})
    if measure.get("dtype") not in {"int", "integer"}:
        return False
    for cell in cells:
        node = graph.get_node(cell)
        value = getattr(node, "value", None)
        if isinstance(value, float) and not value.is_integer():
            return True
    return False


def read_as_numbers(entry: dict[str, Any]) -> None:
    """Declare a measure as float so its numeric cells keep Excel number semantics."""
    measure = entry["structure"]["measure"]
    measure["dtype"] = "float"
    measure.setdefault("bind", {})["read"] = "float"


def reconcile_bindings(
    graph: Any,
    bindings: Mapping[str, Any],
    workbook: Path,
    blank_ranges: Sequence[str] | None = None,
) -> tuple[dict[str, Any], dict[str, str]]:
    """Adjust bindings the catalog would refuse, returning every change.

    Output bindings whose `data_range` has no graph formula cell cover constant
    target cells; those values stay reachable through constant bindings. When
    two formula-direction bindings claim one cell, the `leftover.passthrough`
    claimant is dropped when there is one, otherwise the later binding. Leaf
    cells of the formula closure that no binding names become scalar constant
    bindings so the closure stays fully bound.
    """
    from excel_grapher.grapher.blank_ranges import (
        address_in_blank_ranges,
        normalize_blank_range_specs,
    )
    from excel_grapher.series_bindings.graph_predicates import is_graph_formula_node
    from excel_grapher.series_bindings.ranges import expand_bound_series_addresses_for_graph

    series = list(bindings["series"])
    direction = {
        entry["id"]: next(k for k in ("input", "output", "internal", "constant") if k in entry)
        for entry in series
    }
    table = {entry["id"]: (entry.get("series_context") or {}).get("TABLE") for entry in series}
    cells = {
        entry["id"]: expand_bound_series_addresses_for_graph(graph, entry, workbook=workbook)
        for entry in series
    }
    dropped: dict[str, str] = {}
    for entry in series:
        sid = entry["id"]
        if direction[sid] == "output" and not any(
            is_graph_formula_node(graph, cell) for cell in cells[sid]
        ):
            dropped[sid] = "output binding on constant target cells"
    owners: dict[str, list[str]] = collections.defaultdict(list)
    for entry in series:
        sid = entry["id"]
        if sid in dropped or direction[sid] not in {"internal", "output"}:
            continue
        for cell in cells[sid]:
            owners[cell].append(sid)
    for cell, ids in owners.items():
        live = [sid for sid in ids if sid not in dropped]
        if len(live) < 2:
            continue
        leftovers = [sid for sid in live if table[sid] == "leftover.passthrough"]
        victims = leftovers if leftovers and len(leftovers) < len(live) else live[1:]
        for victim in victims:
            others = [sid for sid in live if sid != victim]
            dropped[victim] = f"formula cell {cell} also bound by {others}"
    kept = [entry for entry in series if entry["id"] not in dropped]
    retyped: dict[str, str] = {}
    for entry in kept:
        sid = entry["id"]
        if direction[sid] in {"input", "constant"} and numeric_text_measure(
            graph, entry, cells[sid]
        ):
            read_as_numbers(entry)
            retyped[sid] = "text measure over numeric cells read as float"
        elif fractional_int_measure(graph, entry, cells[sid]):
            read_as_numbers(entry)
            retyped[sid] = "integer measure over fractional cells read as float"
    bound = {cell for entry in kept for cell in cells[entry["id"]]}
    roots = [
        cell
        for entry in kept
        if direction[entry["id"]] in {"internal", "output"}
        for cell in cells[entry["id"]]
    ]
    seen = set(roots)
    stack = list(roots)
    while stack:
        current = stack.pop()
        for dependency in graph.get_dependencies(current):
            address = str(dependency)
            if address not in seen:
                seen.add(address)
                stack.append(address)
    rects = normalize_blank_range_specs(blank_ranges)
    for address in sorted(seen - bound):
        node = graph.get_node(address)
        if node is None or is_graph_formula_node(graph, address):
            continue
        if address_in_blank_ranges(address, rects):
            continue
        sheet, cell = address.rsplit("!", 1)
        sheet = sheet.strip("'")
        dtype = "string" if isinstance(node.value, str) else "float"
        series_id = "unbound_" + "".join(
            ch.lower() if ch.isalnum() else "_" for ch in f"{sheet}_{cell}"
        )
        kept.append(
            {
                "id": series_id,
                "sheet": sheet,
                "data_range": address,
                "layout": "scalar",
                "constant": {},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": dtype,
                        "bind": {"kind": "data_cell", "read": dtype},
                    },
                    "dimensions": [],
                },
                "key": [],
            }
        )
        dropped[series_id] = f"added constant binding for unbound leaf {address}"
    result = dict(bindings)
    result["series"] = kept
    return result, {**dropped, **retyped}


# ---------------------------------------------------------------------------
# Generation
# ---------------------------------------------------------------------------


def generate_package(
    graph: Any,
    paths: Mapping[str, Path],
    out_dir: Path,
    *,
    reconcile: bool,
    inventory: bool,
    timings: dict[str, float],
) -> dict[str, Any]:
    from excel_grapher.exporter.inverted_tree.deps import bind_blank_rects, reset_blank_rects
    from excel_grapher.exporter.inverted_tree.emit import (
        generate_inverted_tree_modules,
        plan_inverted_tree,
    )
    from excel_grapher.exporter.inverted_tree.named_emit import inventory_named_emission
    from excel_grapher.grapher.blank_ranges import (
        load_blank_ranges_module,
        normalize_blank_range_specs,
    )
    from excel_grapher.series_bindings.load import load_series_bindings

    report: dict[str, Any] = {}
    bindings = load_series_bindings(paths["bindings"])
    if reconcile:
        blank = load_blank_ranges_module(paths["blank_ranges"]) if "blank_ranges" in paths else None
        adjusted, dropped = reconcile_bindings(graph, bindings, paths["workbook"], blank)
        bindings = cast("WorkbookSeriesBindings", adjusted)
        report["binding_adjustments"] = dropped
    blank = load_blank_ranges_module(paths["blank_ranges"]) if "blank_ranges" in paths else None
    if inventory:
        token = bind_blank_rects(normalize_blank_range_specs(blank))
        try:
            started = time.perf_counter()
            catalog, deps, scc_map = plan_inverted_tree(
                graph,
                series_bindings=bindings,
                bindings_workbook=paths["workbook"],
                blank_ranges=blank,
            )
            timings["plan_seconds"] = time.perf_counter() - started
            report["series_counts"] = {
                "formula": len(catalog.formula_series()),
                "output": len(catalog.output_series()),
                "input": len(catalog.input_series()),
                "constant": len(catalog.constant_series()),
                "recurrence_groups": len({scc for scc in scc_map.values() if len(scc) > 1}),
                "formula_cells": sum(len(s.cells) for s in catalog.formula_series()),
            }
            started = time.perf_counter()
            failures = inventory_named_emission(catalog, deps, scc_map, graph)
            timings["inventory_seconds"] = time.perf_counter() - started
            report["lowering_failures"] = failures
            report["lowering_failure_kinds"] = dict(
                collections.Counter(
                    str(item["error"]).split(":")[-1].strip()[:80] for item in failures
                )
            )
            out_dir.mkdir(parents=True, exist_ok=True)
            (out_dir.parent / f"{out_dir.name}.inventory.json").write_text(
                json.dumps(report, indent=2, default=str), encoding="utf-8"
            )
            print(f"lowering failures: {len(failures)}", flush=True)
        finally:
            reset_blank_rects(token)
    started = time.perf_counter()
    modules = generate_inverted_tree_modules(
        graph,
        series_bindings=bindings,
        bindings_workbook=paths["workbook"],
        blank_ranges=blank,
    )
    timings["generate_seconds"] = time.perf_counter() - started
    out_dir.mkdir(parents=True, exist_ok=True)
    for stale in out_dir.glob("*.py"):
        stale.unlink()
    for name, source in modules.items():
        (out_dir / name).write_text(source, encoding="utf-8")
    return report


# ---------------------------------------------------------------------------
# Static structure
# ---------------------------------------------------------------------------


def _segment_bytes(source: str, node: ast.AST) -> int:
    segment = ast.get_source_segment(source, node)
    return len(segment.encode("utf-8")) if segment else 0


def _statement_prefix(node: ast.stmt) -> str:
    if isinstance(node, ast.Assign) and len(node.targets) == 1:
        target = node.targets[0]
        if isinstance(target, ast.Name):
            name = target.id
            for suffix in ("_CELLS", "_DOMAIN", "_REQUIRED", "_SCHEMA", "_DEFAULT", "_LITERALS"):
                if name.endswith(suffix):
                    return suffix
            if name.endswith("_AXIS") or "_AXIS_" in name:
                return "_AXIS"
            return "constant value"
    if isinstance(node, ast.ClassDef):
        return "class"
    if isinstance(node, ast.FunctionDef):
        return "function"
    if isinstance(node, (ast.Import, ast.ImportFrom)):
        return "import"
    return type(node).__name__


def _body_text(source: str, node: ast.FunctionDef) -> str:
    """Source of a function body, without the docstring or the signature."""
    statements = node.body
    if (
        statements
        and isinstance(statements[0], ast.Expr)
        and isinstance(statements[0].value, ast.Constant)
        and isinstance(statements[0].value.value, str)
    ):
        statements = statements[1:]
    if not statements:
        return ""
    lines = source.splitlines(keepends=True)
    first, last = statements[0], statements[-1]
    assert last.end_lineno is not None
    return "".join(lines[first.lineno - 1 : last.end_lineno])


def measure_module(path: Path) -> dict[str, Any]:
    """Attribute a generated module's bytes to statement categories."""
    source = path.read_text(encoding="utf-8")
    tree = ast.parse(source)
    categories: dict[str, int] = collections.Counter()
    counts: dict[str, int] = collections.Counter()
    body_hashes: dict[str, list[str]] = collections.defaultdict(list)
    for node in tree.body:
        categories[_statement_prefix(node)] += _segment_bytes(source, node)
        if isinstance(node, ast.FunctionDef):
            counts["functions"] += 1
            for decorator in node.decorator_list:
                categories["@publish"] += _segment_bytes(source, decorator)
            body = _body_text(source, node).replace(node.name, "<self>")
            body_hashes[hashlib.sha256(body.encode()).hexdigest()].append(node.name)
        if isinstance(node, ast.Assign) and _statement_prefix(node) == "_CELLS":
            value = node.value
            if isinstance(value, ast.Dict):
                counts["provenance_entries"] += len(value.keys)
    for node in ast.walk(tree):
        if isinstance(node, ast.Lambda):
            counts["lambdas"] += 1
        elif isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            if node.func.id in {"lazy_table", "CoordinateReader", "axis_step", "TensorSchema"}:
                counts[node.func.id] += 1
            elif node.func.id == "publish":
                counts["publish"] += 1
        elif isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
            if node.func.attr == "validate":
                counts["schema_validate_calls"] += 1
                categories["schema validate"] += _segment_bytes(source, node)
            elif node.func.attr == "explicit":
                counts["explicit_domains"] += 1
                categories["explicit coordinates"] += _segment_bytes(source, node)
        elif (
            isinstance(node, ast.Tuple)
            and node.elts
            and all(isinstance(item, ast.Tuple) for item in node.elts)
        ):
            counts["coordinate_tuples"] += len(node.elts)
    duplicates = {names[0]: len(names) for names in body_hashes.values() if len(names) > 1}
    counts["duplicate_function_bodies"] = sum(duplicates.values()) - len(duplicates)
    return {
        "bytes": len(source.encode("utf-8")),
        "lines": source.count("\n"),
        "categories": dict(sorted(categories.items(), key=lambda item: -item[1])),
        "counts": dict(counts),
        "duplicate_body_groups": dict(sorted(duplicates.items(), key=lambda item: -item[1])[:20]),
    }


def measure_package(package_dir: Path) -> dict[str, Any]:
    modules: dict[str, Any] = {}
    for path in sorted(package_dir.glob("*.py")):
        modules[path.name] = measure_module(path) | {"sha256": _sha256(path)}
    required = [name for name in modules if name in _GENERATED_MODULES]
    other = [name for name in modules if name not in _GENERATED_MODULES]
    total = sum(modules[name]["bytes"] for name in modules)
    return {
        "modules": modules,
        "source_bytes": total,
        "required_module_bytes": sum(modules[name]["bytes"] for name in required),
        "private_modules": other,
        "under_10_mb": total < 10_000_000,
    }


# ---------------------------------------------------------------------------
# Dynamic measurements
# ---------------------------------------------------------------------------

_IMPORT_PROBE = """
import resource, sys, time
started = time.perf_counter()
import {package}
elapsed = time.perf_counter() - started
usage = resource.getrusage(resource.RUSAGE_SELF)
print(elapsed, usage.ru_maxrss * 1024)
"""

_EVAL_PROBE = """
import inspect, json, resource, sys, time, tracemalloc
use_tracemalloc = {tracemalloc}
if use_tracemalloc:
    tracemalloc.start()
import {package} as package
data = package.data
results = []
names = [name for name in dir(package) if name.startswith("compute_")]
limit = {limit}
if limit:
    names = names[:limit]
for name in names:
    function = getattr(package, name)
    kwargs = {{}}
    for parameter in inspect.signature(function).parameters:
        kwargs[parameter] = getattr(data, parameter.upper() + "_DEFAULT")
    entry = {{"function": name}}
    started = time.perf_counter()
    try:
        result = function(**kwargs)
        entry["seconds"] = time.perf_counter() - started
        entry["cells"] = 1 if function.__domain__ is None else len(function.__domain__)
        samples = []
        for _ in range({warm}):
            started = time.perf_counter()
            function(**kwargs)
            samples.append(time.perf_counter() - started)
        if samples:
            samples.sort()
            entry["warm_median_seconds"] = samples[len(samples) // 2]
    except Exception as exc:
        entry["seconds"] = time.perf_counter() - started
        entry["error"] = f"{{type(exc).__name__}}: {{exc}}"[:300]
    results.append(entry)
usage = resource.getrusage(resource.RUSAGE_SELF)
summary = {{"results": results, "peak_rss_bytes": usage.ru_maxrss * 1024}}
if use_tracemalloc:
    summary["tracemalloc_peak_bytes"] = tracemalloc.get_traced_memory()[1]
print(json.dumps(summary))
"""


def _run_probe(
    package_dir: Path,
    script: str,
    *,
    bytecode: bool,
    standalone: bool,
    timeout: float,
) -> subprocess.CompletedProcess[str]:
    env = dict(os.environ)
    env["PYTHONPATH"] = str(package_dir.parent)
    env.pop("PYTHONSTARTUP", None)
    if not bytecode:
        env["PYTHONDONTWRITEBYTECODE"] = "1"
    prelude = ""
    if standalone:
        prelude = (
            "import sys\n"
            "class _Block:\n"
            "    def find_spec(self, name, path=None, target=None):\n"
            "        if name == 'excel_grapher' or name.startswith('excel_grapher.'):\n"
            "            raise ImportError('excel_grapher is not installed in this process')\n"
            "        return None\n"
            "sys.meta_path.insert(0, _Block())\n"
        )
    command = [sys.executable, "-c", prelude + script]
    if not bytecode:
        command.insert(1, "-B")
    return subprocess.run(
        command, capture_output=True, text=True, env=env, timeout=timeout, check=False
    )


def measure_import(
    package_dir: Path, *, samples: int, standalone: bool, timeout: float
) -> dict[str, Any]:
    package = package_dir.name
    report: dict[str, Any] = {"samples": samples}
    for label, bytecode in (("fresh", False), ("warm", True)):
        elapsed: list[float] = []
        rss: list[float] = []
        errors: list[str] = []
        if bytecode:
            _run_probe(
                package_dir,
                _IMPORT_PROBE.format(package=package),
                bytecode=True,
                standalone=standalone,
                timeout=timeout,
            )
        for _ in range(samples):
            completed = _run_probe(
                package_dir,
                _IMPORT_PROBE.format(package=package),
                bytecode=bytecode,
                standalone=standalone,
                timeout=timeout,
            )
            if completed.returncode != 0:
                errors.append(completed.stderr.strip().splitlines()[-1][:300])
                continue
            seconds, peak = completed.stdout.split()
            elapsed.append(float(seconds))
            rss.append(float(peak))
        entry: dict[str, Any] = {"errors": errors}
        if elapsed:
            entry |= {
                "seconds_median": statistics.median(elapsed),
                "seconds_min": min(elapsed),
                "seconds_max": max(elapsed),
                "peak_rss_bytes_median": statistics.median(rss),
            }
        report[f"{label}_import"] = entry
    for cache in package_dir.glob("__pycache__/*.pyc"):
        report.setdefault("bytecode_bytes", 0)
        report["bytecode_bytes"] += cache.stat().st_size
    return report


def measure_evaluation(
    package_dir: Path,
    *,
    limit: int,
    warm: int,
    standalone: bool,
    use_tracemalloc: bool,
    timeout: float,
) -> dict[str, Any]:
    completed = _run_probe(
        package_dir,
        _EVAL_PROBE.format(
            package=package_dir.name, limit=limit, warm=warm, tracemalloc=use_tracemalloc
        ),
        bytecode=True,
        standalone=standalone,
        timeout=timeout,
    )
    if completed.returncode != 0:
        return {"error": completed.stderr.strip()[-2000:]}
    payload = json.loads(completed.stdout.strip().splitlines()[-1])
    results = payload["results"]
    failures = [item for item in results if "error" in item]
    payload["summary"] = {
        "functions": len(results),
        "failed": len(failures),
        "total_first_call_seconds": sum(item["seconds"] for item in results),
        "cells": sum(item.get("cells", 0) for item in results),
    }
    return payload


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------


def _format_bytes(value: float) -> str:
    return f"{int(value):,}"


def print_summary(report: Mapping[str, Any]) -> None:
    package = report.get("package", {})
    print(f"package: {report.get('package_dir')}")
    print(f"revision: {report['revision']['head'][:12]} ({report['revision']['branch']})")
    for name, module in package.get("modules", {}).items():
        top = ", ".join(
            f"{category} {_format_bytes(size)}"
            for category, size in list(module["categories"].items())[:4]
        )
        print(f"  {name:<14}{_format_bytes(module['bytes']):>14}  {top}")
    print(f"  {'total':<14}{_format_bytes(package.get('source_bytes', 0)):>14}")
    gate = "under" if package.get("under_10_mb") else "OVER"
    print(f"  size gate: {gate} 10,000,000 bytes")
    if "binding_adjustments" in report:
        print(f"  binding adjustments: {len(report['binding_adjustments'])}")
    if "lowering_failures" in report:
        print(f"  lowering failures: {len(report['lowering_failures'])}")
        for kind, count in report.get("lowering_failure_kinds", {}).items():
            print(f"    {count:>5}  {kind}")
    for key, value in report.get("timings", {}).items():
        print(f"  {key}: {value:.3f}")
    imports = report.get("import", {})
    for label in ("fresh_import", "warm_import"):
        entry = imports.get(label)
        if entry and "seconds_median" in entry:
            print(
                f"  {label}: median {entry['seconds_median']:.3f}s "
                f"(min {entry['seconds_min']:.3f}, max {entry['seconds_max']:.3f}), "
                f"peak RSS {_format_bytes(entry['peak_rss_bytes_median'])}"
            )
        elif entry:
            print(f"  {label}: failed {entry['errors'][:1]}")
    evaluation = report.get("evaluation", {})
    if "summary" in evaluation:
        summary = evaluation["summary"]
        print(
            f"  evaluation: {summary['functions']} functions, {summary['failed']} failed, "
            f"{summary['total_first_call_seconds']:.3f}s first calls, "
            f"peak RSS {_format_bytes(evaluation['peak_rss_bytes'])}"
        )
    elif "error" in evaluation:
        print(f"  evaluation failed: {evaluation['error'].splitlines()[-1]}")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    parser.add_argument("sandbox", nargs="?", help="Sandbox directory")
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--bindings", type=Path)
    parser.add_argument("--constraints", type=Path)
    parser.add_argument("--blank-ranges", type=Path)
    parser.add_argument("--targets", type=Path)
    parser.add_argument("--graph", type=Path, help="Saved dependency graph (skips extraction)")
    parser.add_argument("--out-dir", type=Path, help="Directory to write the generated package")
    parser.add_argument("--package-dir", type=Path, help="Measure an existing package")
    parser.add_argument("--report", type=Path, help="Write the JSON report here")
    parser.add_argument("--reconcile-bindings", action="store_true")
    parser.add_argument("--no-inventory", action="store_true", help="Skip the lowering inventory")
    parser.add_argument("--import-samples", type=int, default=3)
    parser.add_argument(
        "--evaluate", type=int, default=0, help="Evaluate the first N outputs (0: all)"
    )
    parser.add_argument("--no-evaluate", action="store_true")
    parser.add_argument("--warm-calls", type=int, default=3)
    parser.add_argument("--tracemalloc", action="store_true")
    parser.add_argument("--standalone", action="store_true", help="Block excel_grapher imports")
    parser.add_argument("--timeout", type=float, default=3600.0)
    args = parser.parse_args(argv)

    warnings.simplefilter("ignore")
    timings: dict[str, float] = {}
    report: dict[str, Any] = {
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "revision": _git_revision(),
        "python": platform.python_version(),
        "platform": platform.platform(),
        "command": [Path(sys.argv[0]).name, *sys.argv[1:]],
    }
    package_dir = args.package_dir
    if args.sandbox is not None and package_dir is None:
        if args.out_dir is None:
            parser.error("--out-dir is required when generating a package")
        paths = _sandbox_paths(args)
        report["inputs"] = _input_hashes(paths)
        graph = _load_graph(paths, timings)
        report["graph"] = {"nodes": sum(1 for _ in graph._nodes)}
        report |= generate_package(
            graph,
            paths,
            args.out_dir,
            reconcile=args.reconcile_bindings,
            inventory=not args.no_inventory,
            timings=timings,
        )
        package_dir = args.out_dir
    if package_dir is None:
        parser.error("give a sandbox directory with --out-dir, or --package-dir")
    package_dir = package_dir.resolve()
    report["package_dir"] = str(package_dir)
    report["package"] = measure_package(package_dir)
    report["import"] = measure_import(
        package_dir,
        samples=args.import_samples,
        standalone=args.standalone,
        timeout=args.timeout,
    )
    if not args.no_evaluate:
        report["evaluation"] = measure_evaluation(
            package_dir,
            limit=args.evaluate,
            warm=args.warm_calls,
            standalone=args.standalone,
            use_tracemalloc=args.tracemalloc,
            timeout=args.timeout,
        )
    report["timings"] = timings
    if args.report is not None:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
    print_summary(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
