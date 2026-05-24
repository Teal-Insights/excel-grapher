"""Helpers for normalizing export targets and emitting Records conversion code."""

from __future__ import annotations

from collections.abc import Sequence

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    normalize_key,
    normalize_range_key,
    parse_address,
    quote_sheet_if_needed,
)
from excel_grapher.exporter.input_groups import NormalizedTargetSpec, TargetShape


def normalize_target_spec(target: str) -> NormalizedTargetSpec:
    from excel_grapher.grapher.target_expansion import split_range_target_on_colon

    split = split_range_target_on_colon(target)
    if split is None:
        sheet, cell = parse_address(normalize_key(target))
        col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
        col = fastpyxl.utils.cell.column_index_from_string(col_str)
        return NormalizedTargetSpec(
            address_or_range=normalize_key(target),
            shape="cell",
            sheet=sheet,
            min_row=int(row),
            min_col=col,
            max_row=int(row),
            max_col=col,
        )

    start_addr, end_addr = split
    sheet, start_a1 = parse_address(start_addr)
    if "!" in end_addr:
        end_sheet, end_a1 = parse_address(end_addr)
        if end_sheet != sheet:
            raise ValueError(f"Range target spans multiple sheets: {target!r}")
    else:
        end_a1 = end_addr

    start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_a1)
    end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_a1)
    c1 = fastpyxl.utils.cell.column_index_from_string(start_col)
    c2 = fastpyxl.utils.cell.column_index_from_string(end_col)
    r1, r2 = (
        (int(start_row), int(end_row))
        if int(start_row) <= int(end_row)
        else (int(end_row), int(start_row))
    )
    if c1 > c2:
        c1, c2 = c2, c1

    row_count = r2 - r1 + 1
    col_count = c2 - c1 + 1
    if row_count == 1 and col_count == 1:
        shape: TargetShape = "cell"
    elif row_count == 1:
        shape = "row_vector"
    elif col_count == 1:
        shape = "col_vector"
    else:
        shape = "rectangle"

    start_col_letter = fastpyxl.utils.cell.get_column_letter(c1)
    end_col_letter = fastpyxl.utils.cell.get_column_letter(c2)
    sheet_q = quote_sheet_if_needed(sheet)
    if r1 == r2 and c1 == c2:
        addr = f"{sheet_q}!{start_col_letter}{r1}"
    else:
        addr = normalize_range_key(f"{sheet_q}!{start_col_letter}{r1}:{end_col_letter}{r2}")

    return NormalizedTargetSpec(
        address_or_range=addr,
        shape=shape,
        sheet=sheet,
        min_row=r1,
        min_col=c1,
        max_row=r2,
        max_col=c2,
    )


def normalize_target_specs(targets: Sequence[str]) -> tuple[NormalizedTargetSpec, ...]:
    return tuple(normalize_target_spec(t) for t in targets)


def record_layout_for_targets(targets: Sequence[str]) -> dict[str, tuple[str, list[str]]]:
    layout: dict[str, tuple[str, list[str]]] = {}
    for target in targets:
        spec = normalize_target_spec(target)
        addresses = list(spec.cell_addresses_row_major())
        if spec.shape == "cell":
            layout[spec.address_or_range] = ("cell", addresses)
        else:
            layout[spec.address_or_range] = (spec.shape, addresses)
    return layout


def emit_records_runtime_helpers() -> list[str]:
    return [
        "def _value_to_records(layout, target, value, *, include_address=True):",
        "    import numpy as np",
        "    kind, addresses = layout[target]",
        "    if kind == 'cell':",
        "        rec = {'value': value}",
        "        if include_address:",
        "            rec['address'] = addresses[0]",
        "        return [rec]",
        "    if not isinstance(value, np.ndarray):",
        "        raise TypeError(f'Expected ndarray for range target {target!r}')",
        "    flat = value.reshape(-1)",
        "    if flat.size != len(addresses):",
        "        raise ValueError('Range size mismatch when building records')",
        "    records = []",
        "    for addr, item in zip(addresses, flat.tolist(), strict=True):",
        "        rec = {'value': item}",
        "        if include_address:",
        "            rec['address'] = addr",
        "        records.append(rec)",
        "    return records",
        "",
        "",
        "def _targets_to_records(ctx, targets, layout, *, include_address=True):",
        "    records = []",
        "    for target, handler in targets.items():",
        "        value = handler(ctx, target)",
        "        records.extend(",
        "            _value_to_records(layout, target, value, include_address=include_address)",
        "        )",
        "    return records",
        "",
        "",
    ]


def emit_record_layout_literal(layout: dict[str, tuple[str, list[str]]]) -> list[str]:
    lines = ["TARGET_RECORD_LAYOUT = {"]
    for target, (kind, addresses) in layout.items():
        addr_literal = ", ".join(repr(a) for a in addresses)
        lines.append(f"    {target!r}: ({kind!r}, [{addr_literal}]),")
    lines.append("}")
    lines.append("")
    return lines
