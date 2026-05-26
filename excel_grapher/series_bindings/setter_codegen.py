"""Generate ``set_*`` functions that apply Records to graph input leaves."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.resolve import resolve_series_bindings
from excel_grapher.series_bindings.types import (
    SeriesResolution,
    WorkbookSeriesBindings,
)

Records = list[dict[str, Any]]


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


def _allowed_record_fields(
    series: dict[str, Any],
    *,
    allow_address: bool,
    requires_address: bool,
) -> set[str]:
    fields = {"value"}
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
) -> list[str]:
    """Emit Python source lines for one series binding setter."""
    if not resolved["leaves"]:
        raise ValueError(f"Cannot codegen setter for {resolved['series_id']!r}: no resolved leaves")
    if not resolved["ok"] and not resolved["requires_address"]:
        raise ValueError(f"Cannot codegen setter for {resolved['series_id']!r}: resolution failed")

    setter = series.get("setter") or {}
    fn_name = str(setter.get("name", f"set_{resolved['series_id']}"))
    strict = bool(setter.get("strict", True))
    allow_address = bool(setter.get("allow_address", False))
    requires_address = bool(resolved["requires_address"])
    key_fields = [str(c) for c in (series.get("key") or [])]
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
    lines.append(f"def {fn_name}(")
    lines.append("    ctx: EvalContext,")
    lines.append("    records: list[dict[str, object]],")
    lines.append("    *,")
    lines.append(f"    strict: bool = {strict!r},")
    lines.append(") -> None:")
    doc = (
        series.get("notes")
        or series.get("sdmx_notes")
        or f"Apply records for {resolved['series_id']}."
    )
    lines.append(f'    """{doc}"""')
    lines.append(f"    key_fields = {tuple(key_fields)!r}")
    lines.append(f"    allow_address = {allow_address!r}")
    lines.append(f"    requires_address = {requires_address!r}")
    lines.append(f"    allowed_fields = {set(allowed)!r}")
    lines.append("    updates: dict[str, object] = {}")
    lines.append("    for index, record in enumerate(records):")
    lines.append("        if strict:")
    lines.append("            unknown = set(record) - allowed_fields")
    lines.append("            if unknown:")
    lines.append(
        '                raise ValueError(f"record[{index}]: unknown fields {sorted(unknown)!r}")'
    )
    lines.append('        if "value" not in record:')
    lines.append(
        "            raise ValueError(f\"record[{index}]: missing required field 'value'\")"
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
    lines.append('        updates[address] = record["value"]')
    lines.append("    if updates:")
    lines.append("        ctx.set_inputs(coerce_inputs_dict(updates))")
    lines.append("")
    return lines


def emit_setters_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
) -> list[str]:
    """Emit all series setter functions for a validated binding manifest."""
    report = resolve_series_bindings(graph, bindings, workbook=workbook)
    lines: list[str] = ["# --- Series binding setters (Records API) ---", ""]
    by_id = {s["id"]: s for s in bindings.get("series", []) if isinstance(s, dict)}
    failed: list[str] = []
    for resolved in report["series"]:
        if not resolved["leaves"]:
            continue
        if not resolved["ok"] and not resolved["requires_address"]:
            failed.append(resolved["series_id"])
            continue
        series = by_id.get(resolved["series_id"])
        if series is None:
            continue
        lines.extend(emit_setter_function(series, resolved))
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
