"""Pass-1 semantic-helper shape contracts (issue #595).

Test-only helpers for Layer A shape-unit suites and the Tiny DSA canary.
These assert the *desired* export surface: binding-named, key-parameterized
helpers instead of per-cell ``cell_*`` dumps. Until collapse lands inside
``CodeGenerator.generate_modules``, callers of these asserts are expected to
fail (RED).
"""

from __future__ import annotations

import ast
import re
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from typing import Any

import fastpyxl

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.evaluator.name_utils import address_to_python_name
from excel_grapher.series_bindings.ranges import expand_data_range
from excel_grapher.series_bindings.setter_codegen import dimension_id_to_param_name


def def_names(source: str) -> set[str]:
    """Return top-level function names defined in *source*."""
    tree = ast.parse(source)
    return {
        node.name for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }


def _function_def(source: str, name: str) -> ast.FunctionDef | ast.AsyncFunctionDef:
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == name:
            return node
    raise AssertionError(f"Expected top-level function {name!r} in source")


def assert_helper_inventory(internals_src: str, expected_ids: Iterable[str]) -> None:
    """Assert every expected series id appears as a top-level ``def``."""
    names = def_names(internals_src)
    missing = sorted(set(expected_ids) - names)
    if missing:
        raise AssertionError(
            "Missing Pass-1 helper defs for bound series ids: "
            f"{missing}; found defs={sorted(names)}"
        )


def assert_helper_signature(
    internals_src: str,
    name: str,
    params: Sequence[str],
) -> None:
    """Assert helper *name* has positional/keyword param names matching *params*.

    Ignores ``*args`` / ``**kwargs``. Annotation text is not compared — only
    names (e.g. ``("ctx", "time_period")``).
    """
    fn = _function_def(internals_src, name)
    args = fn.args
    got: list[str] = [a.arg for a in args.posonlyargs]
    got.extend(a.arg for a in args.args)
    if args.vararg is not None:
        got.append(args.vararg.arg)
    got.extend(a.arg for a in args.kwonlyargs)
    if list(got) != list(params):
        raise AssertionError(
            f"Helper {name!r} signature params {got!r} != expected {list(params)!r}"
        )


def assert_no_cell_defs_for_addresses(
    internals_src: str,
    bound_addresses: Iterable[str],
) -> None:
    """Fail if any bound address still has a ``def cell_*`` emission."""
    names = def_names(internals_src)
    leftovers: list[tuple[str, str]] = []
    for address in bound_addresses:
        normalized = normalize_address(address)
        cell_name = address_to_python_name(normalized)
        if cell_name in names:
            leftovers.append((normalized, cell_name))
    if leftovers:
        preview = ", ".join(f"{addr}->{fn}" for addr, fn in leftovers[:12])
        more = f" (+{len(leftovers) - 12} more)" if len(leftovers) > 12 else ""
        raise AssertionError(
            "Bound formula addresses still have cell_* defs "
            f"(Pass-1 should collapse them): {preview}{more}"
        )


def function_source(source: str, name: str) -> str:
    """Return the source segment for top-level function *name*."""
    fn = _function_def(source, name)
    lines = source.splitlines()
    start = fn.lineno - 1
    end = fn.end_lineno or fn.lineno
    return "\n".join(lines[start:end])


def assert_compute_calls_helper(
    api_src: str,
    compute_name: str,
    helper_name: str,
    *,
    output_addresses: Iterable[str] | None = None,
) -> None:
    """Assert ``compute_*`` invokes *helper_name* instead of xl_cell'ing outputs.

    When *output_addresses* is given, also fail if the compute body still
    contains ``xl_cell(ctx, '<address>')`` for those leaves.
    """
    body = function_source(api_src, compute_name)
    if helper_name not in body:
        raise AssertionError(
            f"{compute_name}() does not reference helper {helper_name!r}; "
            "expected auto-wired output.compute.helper call"
        )
    helper_call = re.compile(rf"\b{re.escape(helper_name)}\s*\(")
    if not helper_call.search(body):
        raise AssertionError(f"{compute_name}() mentions {helper_name!r} but does not call it")
    if output_addresses:
        for address in output_addresses:
            normalized = normalize_address(address)
            needle = f"xl_cell(ctx, '{normalized}')"
            if needle in body:
                raise AssertionError(
                    f"{compute_name}() still xl_cell's output leaf {normalized!r}; "
                    f"expected {helper_name}(...) from record dims"
                )


def series_key_param_names(series: Mapping[str, Any]) -> tuple[str, ...]:
    """Return helper param names for cell-scoped key dimensions (after ``ctx``)."""
    structure = series.get("structure") or {}
    dimensions = structure.get("dimensions") or []
    key_fields = list(series.get("key") or [])
    params: list[str] = []
    for dim in dimensions:
        if not isinstance(dim, Mapping):
            continue
        if dim.get("role") != "key" or dim.get("scope") != "cell":
            continue
        dim_id = str(dim.get("id") or dim.get("concept") or "")
        if not dim_id:
            continue
        concept = dim.get("concept")
        if key_fields and dim_id not in key_fields and concept not in key_fields:
            continue
        params.append(dimension_id_to_param_name(dim_id))
    seen: set[str] = set()
    ordered: list[str] = []
    for param in params:
        if param not in seen:
            seen.add(param)
            ordered.append(param)
    return tuple(ordered)


def expected_helper_signature(series: Mapping[str, Any]) -> tuple[str, ...]:
    """Full helper signature param names including leading ``ctx``."""
    return ("ctx", *series_key_param_names(series))


def _series_has_internal_or_output(series: Mapping[str, Any]) -> bool:
    return "internal" in series or "output" in series


def _cell_has_formula(workbook_path: Path, address: str) -> bool:
    sheet_name, a1 = normalize_address(address).split("!", 1)
    wb = fastpyxl.load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return False
        value = wb[sheet_name][a1].value
    finally:
        wb.close()
    return isinstance(value, str) and value.startswith("=")


def formula_series_ids_from_bindings(
    bindings: Mapping[str, Any],
    *,
    workbook: Path | str | None = None,
    formula_addresses: Iterable[str] | None = None,
) -> set[str]:
    """Return internal/output series ids whose data_range hits formula cells.

    Provide either *formula_addresses* (precomputed graph formula keys) or
    *workbook* (cells inspected via fastpyxl). Constant/input-only series are
    excluded — they are not Pass-1 helper codegen units.
    """
    formula_set: set[str] | None = None
    if formula_addresses is not None:
        formula_set = {normalize_address(a) for a in formula_addresses}

    workbook_path = Path(workbook) if workbook is not None else None
    ids: set[str] = set()
    for series in bindings.get("series", []):
        if not isinstance(series, Mapping):
            continue
        if not _series_has_internal_or_output(series):
            continue
        series_id = series.get("id")
        if not isinstance(series_id, str) or not series_id:
            continue
        data_range = series.get("data_range")
        if not isinstance(data_range, str) or not data_range:
            continue
        try:
            addresses = expand_data_range(
                data_range,
                workbook=workbook_path,
            )
        except ValueError:
            continue
        hit = False
        for address in addresses:
            normalized = normalize_address(address)
            if formula_set is not None:
                if normalized in formula_set:
                    hit = True
                    break
            elif workbook_path is not None:
                if _cell_has_formula(workbook_path, normalized):
                    hit = True
                    break
            else:
                raise ValueError(
                    "formula_series_ids_from_bindings requires workbook= or formula_addresses="
                )
        if hit:
            ids.add(series_id)
    return ids


def bound_formula_addresses_from_bindings(
    bindings: Mapping[str, Any],
    *,
    workbook: Path | str,
    formula_addresses: Iterable[str] | None = None,
) -> list[str]:
    """Addresses in internal/output series data_ranges that are formula cells."""
    workbook_path = Path(workbook)
    formula_set = (
        {normalize_address(a) for a in formula_addresses} if formula_addresses is not None else None
    )
    out: list[str] = []
    seen: set[str] = set()
    for series in bindings.get("series", []):
        if not isinstance(series, Mapping) or not _series_has_internal_or_output(series):
            continue
        data_range = series.get("data_range")
        if not isinstance(data_range, str) or not data_range:
            continue
        try:
            addresses = expand_data_range(data_range, workbook=workbook_path)
        except ValueError:
            continue
        for address in addresses:
            normalized = normalize_address(address)
            if normalized in seen:
                continue
            is_formula = (
                normalized in formula_set
                if formula_set is not None
                else _cell_has_formula(workbook_path, normalized)
            )
            if is_formula:
                seen.add(normalized)
                out.append(normalized)
    return out
