"""Generate ``set_*`` functions that apply Records to graph input leaves."""

from __future__ import annotations

import warnings
from collections.abc import Mapping, Sequence
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


def _allowed_fields_literal(allowed: Sequence[str]) -> str:
    """Emit a deterministic frozenset literal for generated setter calls."""
    inner = ", ".join(repr(field) for field in allowed)
    return f"frozenset({{{inner}}})"


def emit_setter_helpers() -> list[str]:
    """Emit shared private helpers used by generated ``set_*`` functions."""
    return [
        "def _coerce_records(records, measure_field, *, allow_scalar=False) -> Records:",
        "    if not allow_scalar:",
        "        return records",
        "    if not isinstance(records, list):",
        "        if isinstance(records, dict):",
        "            return [records]",
        "        return [{measure_field: records}]",
        "    return records",
        "",
        "def _apply_series_records(",
        "    ctx,",
        "    records,",
        "    *,",
        "    key_fields,",
        "    allowed_fields,",
        "    measure_field,",
        "    leaf_index,",
        "    strict,",
        "    fn_name,",
        "    allow_address=False,",
        "    requires_address=False,",
        ") -> None:",
        "    updates: dict[str, object] = {}",
        "    for index, record in enumerate(records):",
        "        if strict:",
        "            unknown = set(record) - allowed_fields",
        "            if unknown:",
        '                raise ValueError(f"record[{index}]: unknown fields {sorted(unknown)!r}")',
        "        if measure_field not in record:",
        '            raise ValueError(f"record[{index}]: missing required field {measure_field!r}")',
        "        address = None",
        "        if allow_address or requires_address:",
        '            address = record.get("address") or record.get("cell_address")',
        "        if requires_address and address is None:",
        "            raise ValueError(",
        '                f"record[{index}]: address required for {fn_name} (duplicate keys in binding)"',
        "            )",
        "        if address is None:",
        "            if not requires_address:",
        "                missing = [field for field in key_fields if field not in record]",
        "                if missing:",
        '                    raise ValueError(f"record[{index}]: missing key fields {missing!r}")',
        "                key_tuple = tuple((field, record[field]) for field in key_fields)",
        "                address = leaf_index.get(key_tuple)",
        "                if address is None:",
        "                    raise ValueError(",
        '                        f"record[{index}]: no leaf matches key {dict(key_tuple)!r}"',
        "                    )",
        "        updates[address] = record[measure_field]",
        "    if updates:",
        "        ctx.set_inputs(coerce_inputs_dict(updates))",
        "",
    ]


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
        lines.append("    records: Records | Record | Scalar,")
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
    leaf_index_arg = "{}" if requires_address else index_name
    if scalar_shorthand:
        lines.append("    _apply_series_records(")
        lines.append("        ctx,")
        lines.append(f"        _coerce_records(records, {measure_concept!r}, allow_scalar=True),")
    else:
        lines.append("    _apply_series_records(")
        lines.append("        ctx,")
        lines.append("        records,")
    lines.append(f"        key_fields={tuple(key_fields)!r},")
    lines.append(f"        allowed_fields={_allowed_fields_literal(allowed)},")
    lines.append(f"        measure_field={measure_concept!r},")
    lines.append(f"        leaf_index={leaf_index_arg},")
    lines.append("        strict=strict,")
    lines.append(f"        fn_name={fn_name!r},")
    lines.append(f"        allow_address={allow_address!r},")
    lines.append(f"        requires_address={requires_address!r},")
    lines.append("    )")
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
                "Scalar = str | int | float | bool | None",
                "Record = dict[str, object]",
                "Records = list[Record]",
                "",
            ]
        )
    lines.extend(emit_setter_helpers())
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
