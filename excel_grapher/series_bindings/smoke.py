"""Smoke-test generated series-binding setters and compute functions."""

from __future__ import annotations

import importlib
import importlib.util
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import effective_dimension_id
from excel_grapher.series_bindings.resolve import resolve_series_binding
from excel_grapher.series_bindings.setter_codegen import _canonical_key_order
from excel_grapher.series_bindings.types import SeriesResolution, WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import compute_names, setter_names

SetterFn = Callable[[Any, object], None]
ComputeFn = Callable[..., list[dict[str, object]]]


class BindingsSmokeError(RuntimeError):
    """Raised when a generated setter or compute fails a smoke check."""


def _key_fields(series: dict[str, Any]) -> list[str]:
    key = series.get("key")
    if isinstance(key, list):
        return [str(item) for item in key]
    dimensions = (series.get("structure") or {}).get("dimensions") or []
    return [
        effective_dimension_id(dim)
        for dim in dimensions
        if isinstance(dim, dict) and dim.get("role") == "key" and effective_dimension_id(dim)
    ]


def _series_for_setter(bindings: WorkbookSeriesBindings, setter_name: str) -> dict[str, Any]:
    for series in bindings["series"]:
        input_block = series.get("input") or {}
        setter = input_block.get("setter") or series.get("setter")
        if isinstance(setter, dict) and setter.get("name") == setter_name:
            return series
    raise BindingsSmokeError(f"No series declares setter {setter_name!r}")


def _measure_concept(series: dict[str, Any]) -> str:
    measure = (series.get("structure") or {}).get("measure") or {}
    return str(measure.get("concept") or "OBS_VALUE")


def _bump_value(value: object) -> object:
    if isinstance(value, str):
        return f"{value}*"
    if isinstance(value, bool):
        return not value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) + 1.0
    return value


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


def _find_single_key_setter_candidate(
    setter_function_names: list[str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    workbook: Path,
    scheme: dict[str, Any] | None,
) -> tuple[str, dict[str, Any], SeriesResolution] | None:
    for name in setter_function_names:
        series = _series_for_setter(bindings, name)
        if str(series.get("layout") or "series") == "scalar":
            continue
        key_fields = _key_fields(series)
        if len(key_fields) != 1:
            continue
        resolved = resolve_series_binding(
            graph,
            workbook,
            series,
            concept_scheme=scheme,
            direction="input",
        )
        if not resolved["ok"] or resolved["requires_address"] or not resolved["leaves"]:
            continue
        if _canonical_key_order(resolved, key_fields) is None:
            continue
        return name, series, resolved
    return None


def _find_matrix_setter_candidate(
    setter_function_names: list[str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    workbook: Path,
    scheme: dict[str, Any] | None,
) -> tuple[str, dict[str, Any], SeriesResolution] | None:
    for name in setter_function_names:
        series = _series_for_setter(bindings, name)
        if str(series.get("layout") or "series") != "matrix":
            continue
        key_fields = _key_fields(series)
        if len(key_fields) < 2:
            continue
        resolved = resolve_series_binding(
            graph,
            workbook,
            series,
            concept_scheme=scheme,
            direction="input",
        )
        if not resolved["ok"] or resolved["requires_address"] or not resolved["leaves"]:
            continue
        return name, series, resolved
    return None


def _smoke_setter_positional_input(
    pkg: Any,
    setter_name: str,
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    make_context: Callable[[], Any],
) -> None:
    setter = cast(SetterFn, getattr(pkg, setter_name))
    key_fields = _key_fields(series)
    key_field = key_fields[0]
    measure = _measure_concept(series)
    leaves = sorted(resolved["leaves"], key=lambda leaf: leaf["key"][key_field])
    values: list[object] = [leaf["record"][measure] for leaf in leaves]
    target_index = 0
    bumped = _bump_value(values[target_index])
    values[target_index] = bumped
    target_address = leaves[target_index]["address"]
    ctx = make_context()
    setter(ctx, values)
    if ctx.inputs[target_address] != bumped:
        raise BindingsSmokeError(
            f"Setter {setter_name!r} positional input did not update {target_address!r} "
            f"(expected {bumped!r}, got {ctx.inputs[target_address]!r})"
        )


def _smoke_setter_dataframe_input(
    pkg: Any,
    setter_name: str,
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    make_context: Callable[[], Any],
) -> None:
    if importlib.util.find_spec("pandas") is None:
        return
    import pandas as pd

    setter = cast(SetterFn, getattr(pkg, setter_name))
    key_fields = _key_fields(series)
    key_field = key_fields[0]
    measure = _measure_concept(series)
    leaves = sorted(resolved["leaves"], key=lambda leaf: leaf["key"][key_field])
    leaf = leaves[0]
    bumped = _bump_value(leaf["record"][measure])
    frame = pd.DataFrame([{key_field: leaf["key"][key_field], measure: bumped}])
    ctx = make_context()
    setter(ctx, frame)
    if ctx.inputs[leaf["address"]] != bumped:
        raise BindingsSmokeError(
            f"Setter {setter_name!r} DataFrame input did not update {leaf['address']!r} "
            f"(expected {bumped!r}, got {ctx.inputs[leaf['address']]!r})"
        )


def _smoke_setter_matrix_dataframe_input(
    pkg: Any,
    setter_name: str,
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    make_context: Callable[[], Any],
) -> None:
    if importlib.util.find_spec("pandas") is None:
        return
    import pandas as pd

    setter = cast(SetterFn, getattr(pkg, setter_name))
    key_fields = _key_fields(series)
    measure = _measure_concept(series)
    leaf = resolved["leaves"][0]
    bumped = _bump_value(leaf["record"][measure])
    row: dict[str, object] = {field: leaf["key"][field] for field in key_fields}
    row[measure] = bumped
    frame = pd.DataFrame([row])
    ctx = make_context()
    setter(ctx, frame)
    if ctx.inputs[leaf["address"]] != bumped:
        raise BindingsSmokeError(
            f"Setter {setter_name!r} matrix DataFrame input did not update {leaf['address']!r} "
            f"(expected {bumped!r}, got {ctx.inputs[leaf['address']]!r})"
        )


def smoke_test_setters(
    pkg: Any,
    setter_function_names: list[str],
    *,
    bindings: WorkbookSeriesBindings,
    graph: DependencyGraph,
    workbook: Path,
    ctx: Any,
) -> None:
    """Exercise each generated setter with one bumped record.

    Also smoke-tests one single-key series setter with positional values and,
    when pandas is installed, a tidy DataFrame partial update. When a matrix
    setter is available, also smoke-tests one matrix DataFrame partial update.
    """
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
        measure = _measure_concept(series)
        bumped = _bump_value(leaf["record"][measure])
        records: list[dict[str, object]] = [{**key, measure: bumped}]
        setter(ctx, records)
        if ctx.inputs[address] != bumped:
            raise BindingsSmokeError(
                f"Setter {name!r} did not update {address!r} "
                f"(expected {bumped!r}, got {ctx.inputs[address]!r})"
            )

    candidate = _find_single_key_setter_candidate(
        setter_function_names,
        bindings=bindings,
        graph=graph,
        workbook=workbook,
        scheme=scheme,
    )
    if candidate is not None:
        name, series, resolved = candidate
        make_context = cast(Callable[[], Any], pkg.make_context)
        _smoke_setter_positional_input(
            pkg,
            name,
            series,
            resolved,
            make_context=make_context,
        )
        _smoke_setter_dataframe_input(
            pkg,
            name,
            series,
            resolved,
            make_context=make_context,
        )

    matrix_candidate = _find_matrix_setter_candidate(
        setter_function_names,
        bindings=bindings,
        graph=graph,
        workbook=workbook,
        scheme=scheme,
    )
    if matrix_candidate is not None:
        name, series, resolved = matrix_candidate
        make_context = cast(Callable[[], Any], pkg.make_context)
        _smoke_setter_matrix_dataframe_input(
            pkg,
            name,
            series,
            resolved,
            make_context=make_context,
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
        key_fields = _key_fields(series)
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
