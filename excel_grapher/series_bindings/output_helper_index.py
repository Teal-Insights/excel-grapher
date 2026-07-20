"""Address → parameterized output-helper mapping for compute_* codegen.

Mirrors `reader_index` for the output direction: when a published leaf is covered
by a dim-keyed helper (bindings metadata or a post-refactor address overlay),
`compute_*` can call that helper from the static record instead of
`xl_cell(ctx, address)`.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Literal, NotRequired, TypedDict

from excel_grapher.core.address_keys import normalize_key
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import has_output_direction
from excel_grapher.series_bindings.resolve import resolve_series_bindings
from excel_grapher.series_bindings.setter_codegen import dimension_id_to_param_name
from excel_grapher.series_bindings.types import Scalar, WorkbookSeriesBindings

OutputHelperCallMode = Literal["helper", "xl_cell"]
OutputHelperFallbackReason = Literal["unbound"]


class OutputHelperSpec(TypedDict):
    """Declares that one or more addresses are served by a parameterized helper."""

    helper: str
    dims: list[str]


class OutputHelperLeafEntry(TypedDict):
    """One uniquely covered output leaf mapped to its helper call."""

    series_id: str
    helper: str
    dims: list[str]
    keys: dict[str, Scalar]
    kwargs: dict[str, Scalar]
    call_form: str


class OutputHelperIndex(TypedDict):
    """Machine-readable reverse map from Excel addresses to helper call forms."""

    leaves: dict[str, OutputHelperLeafEntry]


class OutputHelperCallResolution(TypedDict):
    """Result of resolving an address to a helper call or `xl_cell` fallback."""

    mode: OutputHelperCallMode
    call_form: str
    reason: str | None
    series_id: str | None
    helper: str | None
    dims: NotRequired[list[str]]
    keys: NotRequired[dict[str, Scalar]]
    kwargs: NotRequired[dict[str, Scalar]]


def format_output_helper_call_form(
    helper: str,
    *,
    dims: list[str],
    record_expr: str = "static_record",
) -> str:
    """Build a compute-body call form that pulls kwargs from a static record.

    Args:
        helper: Parameterized helper function name (`(ctx, **dims) -> CellValue`).
        dims: Binding concept / effective ids whose values become kwargs.
        record_expr: Python expression for the static record dict in generated code.

    Returns:
        Call form such as
        `scaled_output_hot(ctx, time_period=static_record["TIME_PERIOD"])`.
    """
    if not dims:
        return f"{helper}(ctx)"
    parts = ", ".join(
        f"{dimension_id_to_param_name(field)}={record_expr}[{field!r}]" for field in dims
    )
    return f"{helper}(ctx, {parts})"


def helper_spec_from_series(series: Mapping[str, object]) -> OutputHelperSpec | None:
    output = series.get("output")
    if not isinstance(output, dict):
        return None
    compute = output.get("compute")
    if not isinstance(compute, dict):
        return None
    helper = compute.get("helper")
    if not isinstance(helper, dict):
        return None
    name = helper.get("name")
    if not isinstance(name, str) or not name:
        return None
    dims_value = helper.get("dims")
    if dims_value is None:
        key = series.get("key") or []
        dims = [str(field) for field in key] if isinstance(key, list) else []
    elif isinstance(dims_value, list):
        dims = [str(field) for field in dims_value]
    else:
        return None
    return {"helper": name, "dims": dims}


def build_output_helper_index(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    export_addresses: Iterable[str] | None = None,
    address_helpers: Mapping[str, OutputHelperSpec] | None = None,
) -> OutputHelperIndex:
    """Build an address → helper index from bindings and optional overlays.

    Series-level `output.compute.helper` covers every resolved output leaf for
    that series. Optional `address_helpers` overlays win per address (post-refactor
    manifests / partial coverage).
    """
    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="output",
        export_addresses=export_addresses,
    )
    by_id = {
        s["id"]: s
        for s in bindings.get("series", [])
        if isinstance(s, dict) and has_output_direction(s)
    }
    leaves: dict[str, OutputHelperLeafEntry] = {}
    for resolved in report["series"]:
        series = by_id.get(resolved["series_id"])
        if series is None:
            continue
        series_spec = helper_spec_from_series(series)
        for leaf in resolved["leaves"]:
            address = normalize_key(leaf["address"])
            overlay = None if address_helpers is None else address_helpers.get(address)
            if overlay is None and address_helpers is not None:
                # Allow callers to pass un-normalized keys.
                overlay = address_helpers.get(leaf["address"])
            spec = overlay if overlay is not None else series_spec
            if spec is None:
                continue
            key_fields = list(spec["dims"])
            keys = {field: leaf["key"][field] for field in key_fields if field in leaf["key"]}
            # Prefer static record values when key triangulation omitted a field
            # that still appears on the record (e.g. include_in_record dims).
            for field in key_fields:
                if field not in keys and field in leaf["record"]:
                    value = leaf["record"][field]
                    if isinstance(value, (str, int, float, bool)) or value is None:
                        keys[field] = value
            kwargs = {
                dimension_id_to_param_name(field): keys[field]
                for field in key_fields
                if field in keys
            }
            leaves[address] = {
                "series_id": resolved["series_id"],
                "helper": spec["helper"],
                "dims": key_fields,
                "keys": keys,
                "kwargs": kwargs,
                "call_form": format_output_helper_call_form(
                    spec["helper"],
                    dims=key_fields,
                ),
            }
    return {"leaves": leaves}


def resolve_output_helper_ref(
    ref: str,
    *,
    graph: DependencyGraph | None = None,
    bindings: WorkbookSeriesBindings | None = None,
    workbook: Path | str | None = None,
    index: OutputHelperIndex | None = None,
    export_addresses: Iterable[str] | None = None,
    address_helpers: Mapping[str, OutputHelperSpec] | None = None,
) -> OutputHelperCallResolution:
    """Resolve a cell address to a helper call form or `xl_cell` fallback."""
    if index is None:
        if graph is None or bindings is None or workbook is None:
            raise ValueError("resolve_output_helper_ref requires index= or graph/bindings/workbook")
        index = build_output_helper_index(
            graph,
            bindings,
            workbook=workbook,
            export_addresses=export_addresses,
            address_helpers=address_helpers,
        )

    normalized = normalize_key(ref)
    leaf = index["leaves"].get(normalized)
    if leaf is None:
        return {
            "mode": "xl_cell",
            "call_form": f"xl_cell(ctx, {normalized!r})",
            "reason": "unbound",
            "series_id": None,
            "helper": None,
        }
    return {
        "mode": "helper",
        "call_form": leaf["call_form"],
        "reason": None,
        "series_id": leaf["series_id"],
        "helper": leaf["helper"],
        "dims": leaf["dims"],
        "keys": leaf["keys"],
        "kwargs": leaf["kwargs"],
    }


def output_helper_names(index: OutputHelperIndex) -> list[str]:
    """Return sorted unique helper function names referenced by the index."""
    names = {entry["helper"] for entry in index["leaves"].values()}
    return sorted(names)
