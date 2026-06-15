"""Smoke-test generated series-binding setters and compute functions."""

from __future__ import annotations

import importlib
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.resolve import resolve_series_binding
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import compute_names, setter_names

SetterFn = Callable[[Any, list[dict[str, object]]], None]
ComputeFn = Callable[..., list[dict[str, object]]]


class BindingsSmokeError(RuntimeError):
    """Raised when a generated setter or compute fails a smoke check."""


def _key_concepts(series: dict[str, Any]) -> list[str]:
    key = series.get("key")
    if isinstance(key, list):
        return [str(item) for item in key]
    dimensions = (series.get("structure") or {}).get("dimensions") or []
    return [
        str(dim["concept"])
        for dim in dimensions
        if isinstance(dim, dict) and dim.get("role") == "key" and dim.get("concept")
    ]


def _series_for_setter(bindings: WorkbookSeriesBindings, setter_name: str) -> dict[str, Any]:
    for series in bindings["series"]:
        input_block = series.get("input") or {}
        setter = input_block.get("setter") or series.get("setter")
        if isinstance(setter, dict) and setter.get("name") == setter_name:
            return series
    raise BindingsSmokeError(f"No series declares setter {setter_name!r}")


def _series_for_compute(bindings: WorkbookSeriesBindings, compute_name: str) -> dict[str, Any]:
    for series in bindings["series"]:
        output_block = series.get("output") or {}
        compute = output_block.get("compute")
        if isinstance(compute, dict) and compute.get("name") == compute_name:
            return series
    raise BindingsSmokeError(f"No series declares compute {compute_name!r}")


def _import_generated_package(
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


def _cleanup_generated_package(*, package_name: str, output_dir: Path) -> None:
    parent = str(output_dir.parent)
    if parent in sys.path:
        sys.path.remove(parent)
    for name in list(sys.modules):
        if name == package_name or name.startswith(f"{package_name}."):
            sys.modules.pop(name, None)


def smoke_test_setters(
    pkg: Any,
    setter_function_names: list[str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    workbook: Path,
    ctx: Any,
) -> None:
    """Exercise each generated setter with one bumped record."""
    concept_scheme = bindings.get("concept_scheme")
    scheme = concept_scheme if isinstance(concept_scheme, dict) else None
    for name in setter_function_names:
        setter = cast(SetterFn, getattr(pkg, name))
        series = _series_for_setter(bindings, name)
        resolved = resolve_series_binding(
            graph,
            workbook,
            series,
            concept_scheme=scheme,
            direction="input",
        )
        if not resolved["ok"]:
            raise BindingsSmokeError(f"Setter {name!r} resolution failed: {resolved['issues']}")
        if not resolved["leaves"]:
            raise BindingsSmokeError(f"Setter {name!r} resolved no leaves")
        leaf = resolved["leaves"][0]
        key = leaf["key"]
        address = leaf["address"]
        obs_value = cast(int | float, leaf["record"]["OBS_VALUE"])
        records: list[dict[str, object]] = [
            cast(dict[str, object], {**key, "OBS_VALUE": float(obs_value) + 1.0})
        ]
        setter(ctx, records)
        if ctx.inputs[address] != records[0]["OBS_VALUE"]:
            raise BindingsSmokeError(
                f"Setter {name!r} did not update {address!r} "
                f"(expected {records[0]['OBS_VALUE']!r}, got {ctx.inputs[address]!r})"
            )


def smoke_test_computes(
    pkg: Any,
    compute_function_names: list[str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    workbook: Path,
    ctx: Any,
) -> None:
    """Exercise each generated compute function and validate record shape."""
    concept_scheme = bindings.get("concept_scheme")
    scheme = concept_scheme if isinstance(concept_scheme, dict) else None
    for name in compute_function_names:
        compute = cast(ComputeFn, getattr(pkg, name))
        series = _series_for_compute(bindings, name)
        resolved = resolve_series_binding(
            graph,
            workbook,
            series,
            concept_scheme=scheme,
            direction="output",
        )
        if not resolved["ok"]:
            raise BindingsSmokeError(f"Compute {name!r} resolution failed: {resolved['issues']}")
        expected_count = len(resolved["leaves"])
        key_fields = _key_concepts(series)
        records = compute(ctx=ctx)
        if not isinstance(records, list):
            raise BindingsSmokeError(f"Compute {name!r} did not return a list")
        if len(records) != expected_count:
            raise BindingsSmokeError(
                f"Compute {name!r} returned {len(records)} records, expected {expected_count}"
            )
        for record in records:
            if not isinstance(record, dict):
                raise BindingsSmokeError(f"Compute {name!r} returned a non-mapping record")
            if "OBS_VALUE" not in record:
                raise BindingsSmokeError(f"Compute {name!r} record missing OBS_VALUE")
            for field in key_fields:
                if field not in record:
                    raise BindingsSmokeError(f"Compute {name!r} record missing key field {field!r}")


def smoke_test_bindings_module(
    files: dict[str, str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    workbook: Path,
    module_dir: Path,
    package_name: str,
) -> None:
    """Import generated modules and smoke-test all declared setters and computes."""
    pkg = _import_generated_package(files, package_name=package_name, output_dir=module_dir)
    try:
        ctx = pkg.make_context()
        smoke_test_computes(
            pkg,
            compute_names(bindings),
            bindings=bindings,
            graph=graph,
            workbook=workbook,
            ctx=ctx,
        )
        smoke_test_setters(
            pkg,
            setter_names(bindings),
            bindings=bindings,
            graph=graph,
            workbook=workbook,
            ctx=ctx,
        )
    finally:
        _cleanup_generated_package(package_name=package_name, output_dir=module_dir)
