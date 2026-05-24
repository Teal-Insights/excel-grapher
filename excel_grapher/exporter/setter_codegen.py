"""Generate setters.py source from discovered input groups."""

from __future__ import annotations

import re
from collections.abc import Sequence

from excel_grapher.exporter.input_groups import InputGroup, SetterGenerationOptions


def _default_setter_name(group: InputGroup) -> str:
    base = group.group_id
    base = re.sub(r"[^a-z0-9_]+", "_", base.lower()).strip("_")
    if not base:
        base = "group"
    if base[0].isdigit():
        base = f"g_{base}"
    return f"set_{base}"


def _setter_function_name(group: InputGroup, options: SetterGenerationOptions) -> str:
    if options.naming_strategy is not None:
        return options.naming_strategy(group)
    return _default_setter_name(group)


def emit_setters_module(
    groups: Sequence[InputGroup],
    options: SetterGenerationOptions,
) -> str:
    lines: list[str] = [
        "from __future__ import annotations",
        "",
        "from typing import Any",
        "",
        "Scalar = int | float | str | bool | None",
        "Record = dict[str, Scalar | list[str]]",
        "Records = list[Record]",
        "",
        "",
        "def _validate_records(records: Records) -> None:",
        "    for record in records:",
        '        if "value" not in record:',
        "            raise ValueError(\"Record must include 'value'\")",
        "",
        "",
        "def _records_to_inputs(records: Records) -> dict[str, Scalar]:",
        "    _validate_records(records)",
        "    out: dict[str, Scalar] = {}",
        "    for record in records:",
        '        address = record.get("address")',
        "        if not isinstance(address, str):",
        "            raise ValueError(\"Record must include string 'address' for setter application\")",
        '        out[address] = record["value"]',
        "    return out",
        "",
        "",
    ]

    export_names: list[str] = ["_validate_records", "_records_to_inputs"]
    for group in groups:
        fn_name = _setter_function_name(group, options)
        addresses = [c.address for c in group.cells]
        lines.append(f"def {fn_name}(ctx, records: Records) -> None:")
        lines.append(f'    """Apply records to input group {group.group_id!r}."""')
        lines.append("    merged = _records_to_inputs(records)")
        lines.append("    allowed = {")
        for addr in addresses:
            lines.append(f"        {addr!r},")
        lines.append("    }")
        lines.append("    unknown = set(merged) - allowed")
        lines.append("    if unknown:")
        lines.append(
            '        raise ValueError(f"Unknown addresses for setter: {sorted(unknown)!r}")'
        )
        lines.append("    ctx.set_inputs(merged)")
        lines.append("")
        lines.append("")
        export_names.append(fn_name)

    lines.append(f"__all__ = {export_names!r}")
    lines.append("")
    return "\n".join(lines)
