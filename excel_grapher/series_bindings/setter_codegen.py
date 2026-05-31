"""Generate ``set_*`` functions that apply Records to graph input leaves."""

from __future__ import annotations

import warnings
from collections.abc import Mapping
from pathlib import Path
from typing import TYPE_CHECKING, Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.docstrings import (
    emit_docstring_literal,
    resolve_series_function_docstring,
)
from excel_grapher.series_bindings.normalize import has_input_direction
from excel_grapher.series_bindings.resolve import resolve_series_bindings
from excel_grapher.series_bindings.types import (
    SeriesResolution,
    WorkbookSeriesBindings,
)

if TYPE_CHECKING:
    from excel_grapher.series_bindings.docstring_renderers import SeriesDocstringRendererSpec


def _py_literal(value: object) -> str:
    if value is None:
        return "None"
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, str):
        return repr(value)
    if isinstance(value, int) and not isinstance(value, bool):
        return repr(value)
    if isinstance(value, float):
        return repr(value)
    return repr(value)


def _key_tuple_literal(key_fields: list[str], key: Mapping[str, object]) -> str:
    pairs = ", ".join(f"({repr(f)}, {_py_literal(key[f])})" for f in key_fields)
    return f"({pairs},)" if len(key_fields) == 1 else f"({pairs})"


def _measure_concept(series: dict[str, Any]) -> str:
    measure = (series.get("structure") or {}).get("measure") or {}
    return str(measure.get("concept") or "OBS_VALUE")


def _accepts_scalar_shorthand(
    series: dict[str, Any],
    resolved: SeriesResolution,
) -> bool:
    """True when the setter may accept a bare measure value instead of record(s)."""
    if series.get("layout") != "scalar":
        return False
    if series.get("key"):
        return False
    return len(resolved["leaves"]) == 1


def _allowed_record_fields(
    series: dict[str, Any],
    *,
    allow_address: bool,
    requires_address: bool,
) -> set[str]:
    fields = {_measure_concept(series)}
    fields.update(str(c) for c in (series.get("key") or []))
    for dim in (series.get("structure") or {}).get("dimensions") or []:
        if isinstance(dim, dict) and dim.get("include_in_record", True):
            fields.add(str(dim.get("concept", "")))
    for attr in (series.get("structure") or {}).get("attributes") or []:
        if isinstance(attr, dict) and attr.get("include_in_record", False):
            fields.add(str(attr.get("concept", "")))
    fields.update(str(c) for c in (series.get("series_context") or {}))
    if allow_address or requires_address:
        fields.update({"address", "cell_address"})
    return {f for f in fields if f}


def emit_setter_function(
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    graph: DependencyGraph | None = None,
    workbook: Path | str | None = None,
    bindings: WorkbookSeriesBindings | None = None,
    series_docstring_callback: str | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "plain",
) -> list[str]:
    """Emit Python source lines for one series binding setter."""
    if not resolved["leaves"]:
        raise ValueError(f"Cannot codegen setter for {resolved['series_id']!r}: no resolved leaves")
    if not resolved["ok"] and not resolved["requires_address"]:
        raise ValueError(f"Cannot codegen setter for {resolved['series_id']!r}: resolution failed")

    input_block = series.get("input") or {}
    setter = input_block.get("setter") or series.get("setter") or {}
    fn_name = str(setter.get("name", f"set_{resolved['series_id']}"))
    strict = bool(setter.get("strict", True))
    allow_address = bool(setter.get("allow_address", False))
    requires_address = bool(resolved["requires_address"])
    key_fields = [str(c) for c in (series.get("key") or [])]
    measure_concept = _measure_concept(series)
    allowed = sorted(
        _allowed_record_fields(
            series,
            allow_address=allow_address,
            requires_address=requires_address,
        )
    )
    index_name = f"_LEAF_INDEX_{resolved['series_id'].upper()}"

    lines: list[str] = []
    if not requires_address:
        lines.append(f"{index_name} = {{")
        for leaf in resolved["leaves"]:
            key_tuple_src = _key_tuple_literal(key_fields, leaf["key"])
            lines.append(f"    {key_tuple_src}: {repr(leaf['address'])},")
        lines.append("}")
        lines.append("")
    scalar_shorthand = _accepts_scalar_shorthand(series, resolved)
    lines.append(f"def {fn_name}(")
    lines.append("    ctx: EvalContext,")
    if scalar_shorthand:
        lines.append("    records: Records | Record | object,")
    else:
        lines.append("    records: Records,")
    lines.append("    *,")
    lines.append(f"    strict: bool = {strict!r},")
    lines.append(") -> None:")
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
            function_kind="setter",
            function_name=fn_name,
            callback_name=series_docstring_callback,
            docstring_renderer=docstring_renderer,
        )
    else:
        doc = (
            series.get("notes")
            or series.get("sdmx_notes")
            or f"Apply records for {resolved['series_id']}."
        )
    if doc is not None:
        lines.extend(emit_docstring_literal(doc))
    lines.append(f"    key_fields = {tuple(key_fields)!r}")
    lines.append(f"    allow_address = {allow_address!r}")
    lines.append(f"    requires_address = {requires_address!r}")
    lines.append(f"    measure_field = {measure_concept!r}")
    lines.append(f"    allowed_fields = {set(allowed)!r}")
    if scalar_shorthand:
        lines.append("    if not isinstance(records, list):")
        lines.append("        if isinstance(records, dict):")
        lines.append("            records = [records]")
        lines.append("        else:")
        lines.append("            records = [{measure_field: records}]")
    lines.append("    updates: dict[str, object] = {}")
    lines.append("    for index, record in enumerate(records):")
    lines.append("        if strict:")
    lines.append("            unknown = set(record) - allowed_fields")
    lines.append("            if unknown:")
    lines.append(
        '                raise ValueError(f"record[{index}]: unknown fields {sorted(unknown)!r}")'
    )
    lines.append("        if measure_field not in record:")
    lines.append(
        '            raise ValueError(f"record[{index}]: missing required field {measure_field!r}")'
    )
    lines.append("        address = None")
    lines.append("        if allow_address or requires_address:")
    lines.append('            address = record.get("address") or record.get("cell_address")')
    lines.append("        if requires_address and address is None:")
    lines.append(
        "            raise ValueError("
        f'f"record[{{index}}]: address required for {fn_name} (duplicate keys in binding)"'
        ")"
    )
    lines.append("        if address is None:")
    if requires_address:
        lines.append("            pass  # address required; key lookup disabled")
    else:
        lines.append("            missing = [field for field in key_fields if field not in record]")
        lines.append("            if missing:")
        lines.append(
            '                raise ValueError(f"record[{index}]: missing key fields {missing!r}")'
        )
        lines.append(
            "            key_tuple = tuple((field, record[field]) for field in key_fields)"
        )
        lines.append(f"            address = {index_name}.get(key_tuple)")
        lines.append("            if address is None:")
        lines.append(
            '                raise ValueError(f"record[{index}]: no leaf matches key {dict(key_tuple)!r}")'
        )
    lines.append("        updates[address] = record[measure_field]")
    lines.append("    if updates:")
    lines.append("        ctx.set_inputs(coerce_inputs_dict(updates))")
    lines.append("")
    return lines


def emit_setters_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    include_type_aliases: bool = True,
    series_docstring_callback: str | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "plain",
) -> list[str]:
    """Emit all series setter functions for a validated binding manifest."""
    report = resolve_series_bindings(graph, bindings, workbook=workbook, direction="input")
    lines: list[str] = ["# --- Series binding setters (Records API) ---", ""]
    if include_type_aliases:
        lines.extend(
            [
                "Record = dict[str, object]",
                "Records = list[Record]",
                "",
            ]
        )
    by_id = {
        s["id"]: s
        for s in bindings.get("series", [])
        if isinstance(s, dict) and has_input_direction(s)
    }
    failed: list[str] = []
    for resolved in report["series"]:
        if not resolved["ok"] and not resolved["requires_address"]:
            failed.append(resolved["series_id"])
            continue
        if not resolved["leaves"]:
            warnings.warn(
                f"No resolved input cells for series {resolved['series_id']!r}; skipping setter emission",
                UserWarning,
                stacklevel=2,
            )
            continue
        series = by_id.get(resolved["series_id"])
        if series is None:
            continue
        lines.extend(
            emit_setter_function(
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
        raise ValueError(f"Cannot codegen setters: resolution failed for {failed!r}")
    return lines


def generate_setters_module(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
) -> str:
    """Generate a standalone module fragment with setters (requires EvalContext imports)."""
    header = [
        "from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict",
        "",
    ]
    return "\n".join(header + emit_setters_block(graph, workbook, bindings))
