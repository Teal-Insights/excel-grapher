"""Shared harness for ffv2 series-binding validation and codegen smoke tests."""

from __future__ import annotations

import importlib
import sys
from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any, cast

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    bindings_canonical_sha256,
    derive_input_series,
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
    validate_series_bindings,
)
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

SetterFn = Callable[[Any, list[dict[str, object]]], None]
ComputeFn = Callable[..., list[dict[str, object]]]


def setter_names(bindings: WorkbookSeriesBindings) -> list[str]:
    names: list[str] = []
    for series in bindings["series"]:
        input_block = series.get("input") or {}
        setter = input_block.get("setter") or series.get("setter")
        if isinstance(setter, dict) and setter.get("name"):
            names.append(str(setter["name"]))
    return sorted(set(names))


def compute_names(bindings: WorkbookSeriesBindings) -> list[str]:
    names: list[str] = []
    for series in bindings["series"]:
        output_block = series.get("output") or {}
        compute = output_block.get("compute")
        if isinstance(compute, dict) and compute.get("name"):
            names.append(str(compute["name"]))
    return sorted(set(names))


def all_series_targets(
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path,
) -> list[str]:
    targets: list[str] = []
    for series in bindings["series"]:
        data_range = series.get("data_range")
        if isinstance(data_range, str):
            targets.extend(expand_data_range(data_range, workbook=workbook))
    return sorted(set(targets))


def validate_ffv2_bindings(
    workbook: Path,
    bindings: WorkbookSeriesBindings,
) -> dict[str, Any]:
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    return {
        "graph": graph,
        "targets": targets,
        "report": report,
        "canonical_sha256": bindings_canonical_sha256(bindings),
        "input_series": derive_input_series(graph, bindings, workbook=workbook),
    }


def import_generated_package(
    files: dict[str, str],
    *,
    package_name: str,
    output_dir: Path,
) -> Any:
    output_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in files.items():
        (output_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(output_dir.parent))
    try:
        for name in list(sys.modules):
            if name == package_name or name.startswith(f"{package_name}."):
                sys.modules.pop(name, None)
        return importlib.import_module(package_name)
    except Exception:
        sys.path.remove(str(output_dir.parent))
        raise


def cleanup_generated_package(*, package_name: str, output_dir: Path) -> None:
    parent = str(output_dir.parent)
    if parent in sys.path:
        sys.path.remove(parent)
    for name in list(sys.modules):
        if name == package_name or name.startswith(f"{package_name}."):
            sys.modules.pop(name, None)


def generate_modules_package(
    graph: Any,
    *,
    targets: list[str],
    bindings: WorkbookSeriesBindings,
    workbook: Path,
) -> dict[str, str]:
    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def assert_compute_functions(
    pkg: Any,
    compute_function_names: list[str],
    *,
    ctx: Any,
    expected_period_count: int = 16,
) -> None:
    for name in compute_function_names:
        compute = cast(ComputeFn, getattr(pkg, name))
        records = compute(ctx=ctx)
        assert isinstance(records, list)
        assert len(records) == expected_period_count, name
        for record in records:
            assert "OBS_VALUE" in record
            assert "TIME_PERIOD" in record
            assert isinstance(record["TIME_PERIOD"], datetime)


def assert_setter_functions(
    pkg: Any,
    setter_function_names: list[str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: Any,
    workbook: Path,
    ctx: Any,
) -> None:
    concept_scheme = bindings.get("concept_scheme")
    for name in setter_function_names:
        setter = cast(SetterFn, getattr(pkg, name))
        series = _series_for_setter(bindings, name)
        resolved = resolve_series_binding(
            graph,
            workbook,
            series,
            concept_scheme=concept_scheme if isinstance(concept_scheme, dict) else None,
            direction="input",
        )
        assert resolved["ok"] is True, name
        leaf = resolved["leaves"][0]
        key = leaf["key"]
        address = leaf["address"]
        obs_value = cast(int | float, leaf["record"]["OBS_VALUE"])
        records: list[dict[str, object]] = [
            cast(dict[str, object], {**key, "OBS_VALUE": float(obs_value) + 1.0})
        ]
        setter(ctx, records)
        assert ctx.inputs[address] == records[0]["OBS_VALUE"], name


def _series_for_setter(bindings: WorkbookSeriesBindings, setter_name: str) -> dict[str, Any]:
    for series in bindings["series"]:
        input_block = series.get("input") or {}
        setter = input_block.get("setter") or series.get("setter")
        if isinstance(setter, dict) and setter.get("name") == setter_name:
            return series
    raise KeyError(f"No series declares setter {setter_name!r}")


def run_ffv2_binding_checks(
    workbook: Path,
    bindings_path: Path,
    *,
    module_dir: Path,
    package_name: str = "ffv2_module",
) -> dict[str, Any]:
    bindings = load_series_bindings(bindings_path)
    validation = validate_ffv2_bindings(workbook, bindings)
    report = validation["report"]
    if not report["ok"]:
        errors = [issue for issue in report["issues"] if issue["level"] == "error"]
        raise AssertionError(f"Binding validation failed: {errors}")

    files = generate_modules_package(
        validation["graph"],
        targets=validation["targets"],
        bindings=bindings,
        workbook=workbook,
    )
    pkg = import_generated_package(files, package_name=package_name, output_dir=module_dir)
    try:
        ctx = pkg.make_context()
        assert_compute_functions(pkg, compute_names(bindings), ctx=ctx)
        assert_setter_functions(
            pkg,
            setter_names(bindings),
            bindings=bindings,
            graph=validation["graph"],
            workbook=workbook,
            ctx=ctx,
        )
    finally:
        cleanup_generated_package(package_name=package_name, output_dir=module_dir)

    return {
        "bindings": bindings,
        "validation": validation,
        "generated_files": files,
        "setters": setter_names(bindings),
        "computes": compute_names(bindings),
    }
