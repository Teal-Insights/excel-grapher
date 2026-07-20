"""Generate `compute_*` functions that return Records from graph output cells."""

from __future__ import annotations

import warnings
from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import TYPE_CHECKING, Any

from excel_grapher.core.address_keys import normalize_key
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.codegen_literals import (
    emit_compute_preamble_lines,
    py_scalar_literal,
    resolution_includes_datetime,
    resolutions_include_datetime,
)
from excel_grapher.series_bindings.docstrings import (
    emit_docstring_literal,
    resolve_series_function_docstring,
)
from excel_grapher.series_bindings.normalize import has_output_direction
from excel_grapher.series_bindings.output_helper_index import (
    OutputHelperIndex,
    OutputHelperLeafEntry,
    OutputHelperSpec,
    build_output_helper_index,
    format_output_helper_call_form,
)
from excel_grapher.series_bindings.resolve import (
    resolve_series_bindings,
    warn_series_resolution_issues,
)
from excel_grapher.series_bindings.types import (
    SeriesResolution,
    WorkbookSeriesBindings,
)

if TYPE_CHECKING:
    from excel_grapher.series_bindings.docstring_renderers import SeriesDocstringRendererSpec
    from excel_grapher.series_bindings.docstrings import SeriesBindingDocstringCallbackSpec


def _record_literal(record: dict[str, object]) -> str:
    items = ", ".join(f"{repr(k)}: {py_scalar_literal(v)}" for k, v in sorted(record.items()))
    return f"{{{items}}}"


def _measure_concept(series: dict[str, Any]) -> str:
    measure = (series.get("structure") or {}).get("measure") or {}
    return str(measure.get("concept") or "OBS_VALUE")


def _leaves_table_name(series_id: str) -> str:
    return f"_OUTPUT_LEAVES_{series_id.upper()}"


def _helper_entry_for_leaf(
    leaf_address: str,
    *,
    helper_index: OutputHelperIndex | None,
) -> OutputHelperLeafEntry | None:
    if helper_index is None:
        return None
    return helper_index["leaves"].get(normalize_key(leaf_address))


def _series_helper_coverage(
    resolved: SeriesResolution,
    *,
    helper_index: OutputHelperIndex | None,
) -> tuple[str, list[str]] | None:
    """Return `(helper, dims)` when every leaf shares one helper mapping."""
    if helper_index is None or not resolved["leaves"]:
        return None
    entries: list[OutputHelperLeafEntry] = []
    for leaf in resolved["leaves"]:
        entry = _helper_entry_for_leaf(leaf["address"], helper_index=helper_index)
        if entry is None:
            return None
        entries.append(entry)
    first = entries[0]
    helper = first["helper"]
    dims = list(first["dims"])
    for entry in entries[1:]:
        if entry["helper"] != helper or list(entry["dims"]) != dims:
            return None
    return helper, dims


def _emit_measure_assignment_lines(
    *,
    homogeneous: tuple[str, list[str]] | None,
    indent: str = "            ",
) -> list[str]:
    """Emit the try-body lines that assign `record[measure_field]`."""
    if homogeneous is not None:
        helper, dims = homogeneous
        call = format_output_helper_call_form(helper, dims=dims)
        return [f"{indent}record[measure_field] = {call}"]
    return [f"{indent}record[measure_field] = xl_cell(ctx, address)"]


def _emit_mixed_measure_assignment_lines(
    resolved: SeriesResolution,
    *,
    helper_index: OutputHelperIndex,
    indent: str = "            ",
) -> list[str]:
    """Emit per-leaf branching when helper coverage is partial or heterogeneous."""
    lines: list[str] = []
    # Group addresses by (helper, dims) for compact elif chains.
    groups: dict[tuple[str, tuple[str, ...]], list[str]] = {}
    uncovered: list[str] = []
    for leaf in resolved["leaves"]:
        address = normalize_key(leaf["address"])
        entry = helper_index["leaves"].get(address)
        if entry is None:
            uncovered.append(address)
            continue
        key = (entry["helper"], tuple(entry["dims"]))
        groups.setdefault(key, []).append(address)

    first = True
    for (helper, dims), addresses in groups.items():
        call = format_output_helper_call_form(helper, dims=list(dims))
        if len(addresses) == 1:
            cond = f"address == {addresses[0]!r}"
        else:
            addr_set = ", ".join(repr(addr) for addr in addresses)
            cond = f"address in {{{addr_set}}}"
        keyword = "if" if first else "elif"
        lines.append(f"{indent}{keyword} {cond}:")
        lines.append(f"{indent}    record[measure_field] = {call}")
        first = False
    if uncovered or first:
        # Always keep an xl_cell path for uncovered leaves (and as else).
        if first:
            lines.append(f"{indent}record[measure_field] = xl_cell(ctx, address)")
        else:
            lines.append(f"{indent}else:")
            lines.append(f"{indent}    record[measure_field] = xl_cell(ctx, address)")
    return lines


def emit_output_leaves_table(
    series: dict[str, Any],
    resolved: SeriesResolution,
) -> list[str]:
    """Emit the `_OUTPUT_LEAVES_*` table for one resolved output series."""
    if not resolved["leaves"]:
        raise ValueError(
            f"Cannot codegen leaves for {resolved['series_id']!r}: no resolved output cells"
        )
    measure_concept = _measure_concept(series)
    leaves_name = _leaves_table_name(resolved["series_id"])
    lines: list[str] = [f"{leaves_name}: list[tuple[str, Record]] = ["]
    for leaf in resolved["leaves"]:
        static_record: dict[str, object] = {
            str(k): v for k, v in leaf["record"].items() if k != measure_concept
        }
        lines.append(f"    ({repr(leaf['address'])}, {_record_literal(static_record)}),")
    lines.append("]")
    lines.append("")
    return lines


def emit_output_leaves_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
) -> list[str]:
    """Emit all `_OUTPUT_LEAVES_*` tables for output series (modular export)."""
    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="output",
        export_addresses=export_addresses,
    )
    lines: list[str] = [
        "# --- Series binding output leaf tables ---",
        "",
    ]
    if resolutions_include_datetime(report["series"]):
        lines.extend(["import datetime", ""])
    lines.extend(
        [
            "Record = dict[str, object]",
            "",
        ]
    )
    by_id = {
        s["id"]: s
        for s in bindings.get("series", [])
        if isinstance(s, dict) and has_output_direction(s)
    }
    failed: list[str] = []
    emitted = False
    for resolved in report["series"]:
        if not resolved["ok"]:
            failed.append(resolved["series_id"])
            continue
        if not resolved["leaves"]:
            continue
        series = by_id.get(resolved["series_id"])
        if series is None:
            continue
        lines.extend(emit_output_leaves_table(series, resolved))
        emitted = True
    if failed:
        raise ValueError(f"Cannot codegen output leaf tables: resolution failed for {failed!r}")
    if not emitted:
        return ["# --- Series binding output leaf tables ---", ""]
    return lines


def emit_compute_function(
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    graph: DependencyGraph | None = None,
    workbook: Path | str | None = None,
    bindings: WorkbookSeriesBindings | None = None,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
    include_datetime_import: bool = True,
    include_leaves_table: bool = True,
    helper_index: OutputHelperIndex | None = None,
) -> list[str]:
    """Emit Python source lines for one series binding output compute function.

    When `helper_index` (or series `output.compute.helper`) covers a leaf, the
    measure is evaluated via the parameterized helper keyed by static record dims.
    Uncovered leaves keep the `xl_cell(ctx, address)` path.

    Generated measure evaluation catches `XlErrorException`, stores the `XlError`
    code on the measure field, and continues the series. Non-Excel exceptions still
    propagate.
    """
    if not resolved["leaves"]:
        raise ValueError(
            f"Cannot codegen compute for {resolved['series_id']!r}: no resolved output cells"
        )
    if not resolved["ok"]:
        raise ValueError(f"Cannot codegen compute for {resolved['series_id']!r}: resolution failed")

    effective_index: OutputHelperIndex | None = helper_index
    if (
        effective_index is None
        and graph is not None
        and workbook is not None
        and bindings is not None
    ):
        effective_index = build_output_helper_index(graph, bindings, workbook=workbook)
    elif effective_index is None:
        # Build a tiny index from this series alone when bindings context is absent.
        from excel_grapher.series_bindings.output_helper_index import helper_spec_from_series

        spec = helper_spec_from_series(series)
        if spec is not None:
            leaves: dict[str, OutputHelperLeafEntry] = {}
            for leaf in resolved["leaves"]:
                address = normalize_key(leaf["address"])
                leaves[address] = {
                    "series_id": resolved["series_id"],
                    "helper": spec["helper"],
                    "dims": list(spec["dims"]),
                    "keys": {
                        field: leaf["key"][field] for field in spec["dims"] if field in leaf["key"]
                    },
                    "kwargs": {},
                    "call_form": format_output_helper_call_form(
                        spec["helper"], dims=list(spec["dims"])
                    ),
                }
            effective_index = OutputHelperIndex(leaves=leaves)

    output = series.get("output") or {}
    compute = output.get("compute") or {}
    fn_name = str(compute.get("name", f"compute_{resolved['series_id']}"))
    include_address = bool(compute.get("include_address", False))
    leaves_name = _leaves_table_name(resolved["series_id"])
    homogeneous = _series_helper_coverage(resolved, helper_index=effective_index)

    lines: list[str] = []
    if include_datetime_import and resolution_includes_datetime(resolved):
        lines.extend(["import datetime", ""])
    if include_leaves_table:
        lines.extend(emit_output_leaves_table(series, resolved))
    lines.append(f"def {fn_name}(ctx=None, *, inputs=None) -> Records:")
    if series_docstring_callback is not None and (
        graph is None or workbook is None or bindings is None
    ):
        raise ValueError("series_docstring_callback requires graph, workbook, and bindings context")
    if graph is not None and workbook is not None and bindings is not None:
        doc = resolve_series_function_docstring(
            graph=graph,
            workbook=workbook,
            bindings=bindings,
            series=series,
            resolution=resolved,
            function_kind="compute",
            function_name=fn_name,
            callback_spec=series_docstring_callback,
            docstring_renderer=docstring_renderer,
        )
    else:
        doc = (
            series.get("notes")
            or series.get("sdmx_notes")
            or f"Compute records for {resolved['series_id']}."
        )
    if doc is not None:
        lines.extend(emit_docstring_literal(doc))
    lines.append("    if ctx is None:")
    lines.append("        ctx = make_context(inputs)")
    lines.append("    elif inputs is not None:")
    lines.append(
        "        warnings.warn("
        '"inputs will be ignored because ctx was provided", '
        "UserWarning, stacklevel=2)"
    )
    lines.append(f"    measure_field = {_measure_concept(series)!r}")
    lines.append(f"    include_address = {include_address!r}")
    lines.append("    records: Records = []")
    lines.append(f"    for address, static_record in {leaves_name}:")
    lines.append("        record = dict(static_record)")
    lines.append("        try:")
    if homogeneous is not None:
        lines.extend(_emit_measure_assignment_lines(homogeneous=homogeneous))
    elif effective_index is not None and any(
        _helper_entry_for_leaf(leaf["address"], helper_index=effective_index) is not None
        for leaf in resolved["leaves"]
    ):
        lines.extend(
            _emit_mixed_measure_assignment_lines(
                resolved,
                helper_index=effective_index,
            )
        )
    else:
        lines.extend(_emit_measure_assignment_lines(homogeneous=None))
    lines.append("        except XlErrorException as err:")
    lines.append("            record[measure_field] = err.code")
    lines.append("        if include_address:")
    lines.append('            record["address"] = address')
    lines.append("        records.append(record)")
    lines.append("    return records")
    lines.append("")
    return lines


def emit_computes_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
    include_type_aliases: bool = True,
    include_leaves_tables: bool = True,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
    helper_index: OutputHelperIndex | None = None,
    address_helpers: Mapping[str, OutputHelperSpec] | None = None,
) -> list[str]:
    """Emit all series output compute functions for a validated binding manifest."""
    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="output",
        export_addresses=export_addresses,
    )
    effective_index = helper_index
    if effective_index is None:
        effective_index = build_output_helper_index(
            graph,
            bindings,
            workbook=workbook,
            export_addresses=export_addresses,
            address_helpers=address_helpers,
        )
    lines: list[str] = ["# --- Series binding output compute (Records API) ---", ""]
    include_datetime = resolutions_include_datetime(report["series"])
    if include_type_aliases:
        lines.extend(emit_compute_preamble_lines(include_datetime=include_datetime))
    datetime_import_done = include_type_aliases and include_datetime
    by_id = {
        s["id"]: s
        for s in bindings.get("series", [])
        if isinstance(s, dict) and has_output_direction(s)
    }
    failed: list[str] = []
    for resolved in report["series"]:
        if not resolved["ok"]:
            failed.append(resolved["series_id"])
            continue
        if not resolved["leaves"]:
            warnings.warn(
                f"No resolved output cells for series {resolved['series_id']!r}; skipping compute emission",
                UserWarning,
                stacklevel=2,
            )
            continue
        warn_series_resolution_issues(resolved)
        series = by_id.get(resolved["series_id"])
        if series is None:
            continue
        fn_include_datetime_import = (
            resolution_includes_datetime(resolved) and not datetime_import_done
        )
        lines.extend(
            emit_compute_function(
                series,
                resolved,
                graph=graph,
                workbook=workbook,
                bindings=bindings,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
                include_datetime_import=fn_include_datetime_import,
                include_leaves_table=include_leaves_tables,
                helper_index=effective_index,
            )
        )
        if fn_include_datetime_import:
            datetime_import_done = True
    if failed:
        raise ValueError(
            f"Cannot codegen output compute functions: resolution failed for {failed!r}"
        )
    return lines


def generate_computes_module(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
) -> str:
    """Generate a standalone module fragment with output compute functions."""
    header = [
        "import warnings",
        "",
        "from excel_grapher.core.types import XlErrorException",
        "from excel_grapher.runtime.cache import EvalContext, xl_cell",
        "",
    ]
    return "\n".join(header + emit_computes_block(graph, workbook, bindings))
