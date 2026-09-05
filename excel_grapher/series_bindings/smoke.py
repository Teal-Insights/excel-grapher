"""Smoke-test generated series-binding compute functions."""

from __future__ import annotations

import importlib
import inspect
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.input_coerce import input_value_map_from_series
from excel_grapher.series_bindings.resolve import resolve_series_binding
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import compute_names


class BindingsSmokeError(RuntimeError):
    """Raised when a generated compute fails a smoke check."""


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


def smoke_test_computes(
    pkg: Any,
    compute_function_names: list[str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    workbook: Path,
) -> None:
    """Call each inverted-tree `compute_*` with workbook-default input leaves."""
    data = importlib.import_module(f"{pkg.__name__}.data")
    concept_scheme = bindings.get("concept_scheme")
    scheme = concept_scheme if isinstance(concept_scheme, dict) else None
    for name in compute_function_names:
        compute = getattr(pkg, name)
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
        kwargs = _inverted_tree_default_kwargs(compute, data, bindings)
        try:
            result = compute(**kwargs)
        except Exception as exc:
            raise BindingsSmokeError(
                f"Compute {name!r} raised {type(exc).__name__}: {exc}"
            ) from exc
        if not isinstance(result, tuple):
            raise BindingsSmokeError(
                f"Compute {name!r} did not return a tuple (got {type(result).__name__})"
            )
        if len(result) != expected_count:
            raise BindingsSmokeError(
                f"Compute {name!r} returned {len(result)} values, expected {expected_count}"
            )


def _public_compute_arg(series: dict[str, Any], workbook_value: object) -> object:
    """Invert `input.value_map` so smoke calls use public keys, not needles."""
    mapping = input_value_map_from_series(series)
    if mapping is None:
        return workbook_value
    for key, needle in mapping.items():
        if needle == workbook_value:
            return key
    return workbook_value


def _inverted_tree_default_kwargs(
    compute: Callable[..., Any],
    data: Any,
    bindings: WorkbookSeriesBindings,
) -> dict[str, Any]:
    """Bind required compute parameters from `data.py` `*_DEFAULT` leaves."""
    series_by_id = {
        str(series["id"]): series
        for series in bindings.get("series", [])
        if isinstance(series, dict) and series.get("id") is not None
    }
    kwargs: dict[str, Any] = {}
    for name, param in inspect.signature(compute).parameters.items():
        if param.kind in {inspect.Parameter.VAR_POSITIONAL, inspect.Parameter.VAR_KEYWORD}:
            continue
        if param.default is not inspect.Parameter.empty:
            continue
        default_name = f"{name.upper()}_DEFAULT"
        if not hasattr(data, default_name):
            compute_name = getattr(compute, "__name__", "compute")
            raise BindingsSmokeError(
                f"Compute {compute_name!r} required argument {name!r} has no "
                f"{default_name} in data.py"
            )
        workbook_value = getattr(data, default_name)
        series = series_by_id.get(name)
        kwargs[name] = (
            _public_compute_arg(series, workbook_value) if series is not None else workbook_value
        )
    return kwargs


def smoke_test_bindings_module(
    files: dict[str, str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    workbook: Path,
    module_dir: Path,
    package_name: str,
) -> None:
    """Import generated modules and smoke-test public compute functions."""
    pkg = _import_generated_package(files, package_name=package_name, output_dir=module_dir)
    try:
        smoke_test_computes(
            pkg,
            compute_names(bindings),
            bindings=bindings,
            graph=graph,
            workbook=workbook,
        )
    finally:
        _cleanup_generated_package(package_name=package_name, output_dir=module_dir)
