"""Generate `compute_*` functions that return Records from graph output cells."""

from __future__ import annotations

import warnings
from collections.abc import Iterable
from pathlib import Path
from typing import TYPE_CHECKING, Any

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


def emit_compute_function(
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    graph: DependencyGraph | None = None,
    workbook: Path | str | None = None,
    bindings: WorkbookSeriesBindings | None = None,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit Python source lines for one series binding output compute function."""
    if not resolved["leaves"]:
        raise ValueError(
            f"Cannot codegen compute for {resolved['series_id']!r}: no resolved output cells"
        )
    if not resolved["ok"]:
        raise ValueError(f"Cannot codegen compute for {resolved['series_id']!r}: resolution failed")

    output = series.get("output") or {}
    compute = output.get("compute") or {}
    fn_name = str(compute.get("name", f"compute_{resolved['series_id']}"))
    include_address = bool(compute.get("include_address", False))
    measure_concept = _measure_concept(series)
    leaves_name = f"_OUTPUT_LEAVES_{resolved['series_id'].upper()}"

    lines: list[str] = []
    if resolution_includes_datetime(resolved):
        lines.extend(["import datetime", ""])
    lines.append(f"{leaves_name} = [")
    for leaf in resolved["leaves"]:
        static_record: dict[str, object] = {
            str(k): v for k, v in leaf["record"].items() if k != measure_concept
        }
        lines.append(f"    ({repr(leaf['address'])}, {_record_literal(static_record)}),")
    lines.append("]")
    lines.append("")
    lines.append(f"def {fn_name}(inputs=None, *, ctx=None) -> Records:")
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
    lines.append(f"    measure_field = {measure_concept!r}")
    lines.append(f"    include_address = {include_address!r}")
    lines.append("    records: Records = []")
    lines.append(f"    for address, static_record in {leaves_name}:")
    lines.append("        record = dict(static_record)")
    lines.append("        record[measure_field] = xl_cell(ctx, address)")
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
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit all series output compute functions for a validated binding manifest."""
    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="output",
        export_addresses=export_addresses,
    )
    lines: list[str] = ["# --- Series binding output compute (Records API) ---", ""]
    if include_type_aliases:
        include_datetime = resolutions_include_datetime(report["series"])
        lines.extend(emit_compute_preamble_lines(include_datetime=include_datetime))
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
        lines.extend(
            emit_compute_function(
                series,
                resolved,
                graph=graph,
                workbook=workbook,
                bindings=bindings,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
            )
        )
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
        "from excel_grapher.runtime.cache import EvalContext, xl_cell",
        "",
    ]
    return "\n".join(header + emit_computes_block(graph, workbook, bindings))
