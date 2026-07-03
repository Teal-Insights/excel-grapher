"""Configuration for TACO index construction."""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_cell_key, parse_address
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey

if TYPE_CHECKING:
    from excel_grapher.series_bindings.types import WorkbookSeriesBindings


@dataclass(frozen=True, slots=True)
class TacoBuildConfig:
    """Controls which cells may participate in TACO range-pattern compression.

    For codegen, start with ``exclude_targets`` and ``exclude_input_keys`` so
    boundary cells stay at cell granularity. Use ``internal_only`` to compress
    only formula nodes that are neither targets nor declared inputs.
    """

    exclude_targets: bool = False
    exclude_input_keys: frozenset[NodeKey] = frozenset()
    internal_only: bool = False

    @classmethod
    def for_codegen(
        cls,
        graph: DependencyGraph | None = None,
        *,
        input_keys: frozenset[NodeKey] | None = None,
        internal_only: bool = True,
    ) -> TacoBuildConfig:
        """Preset for codegen: keep targets and inputs uncompressed."""
        if input_keys is None and graph is not None:
            input_keys = input_keys_from_graph(graph)
        return cls(
            exclude_targets=True,
            exclude_input_keys=input_keys or frozenset(),
            internal_only=internal_only,
        )

    @classmethod
    def for_codegen_export(
        cls,
        graph: DependencyGraph,
        *,
        input_ranges: Sequence[str] | None = None,
        series_bindings: WorkbookSeriesBindings | None = None,
        bindings_workbook: Path | str | None = None,
        export_addresses: Iterable[str] | None = None,
        internal_only: bool = True,
    ) -> TacoBuildConfig:
        """Preset for codegen export: exclude targets, inputs, and setter leaves."""
        input_keys = codegen_boundary_keys(
            graph,
            input_ranges=input_ranges,
            series_bindings=series_bindings,
            bindings_workbook=bindings_workbook,
            export_addresses=export_addresses,
        )
        return cls.for_codegen(input_keys=input_keys, internal_only=internal_only)


def input_keys_from_graph(graph: DependencyGraph) -> frozenset[NodeKey]:
    """Return sheet-qualified keys marked ``input`` in ``graph.leaf_classification``."""
    lc = graph.leaf_classification
    if not lc:
        return frozenset()
    return frozenset(k for k, role in lc.items() if role == "input")


def input_keys_from_ranges(input_ranges: Sequence[str] | None) -> frozenset[NodeKey]:
    """Expand declared ``input_ranges`` to sheet-qualified cell keys."""
    if not input_ranges:
        return frozenset()
    keys: set[NodeKey] = set()
    for range_str in input_ranges:
        keys.update(_expand_sheet_qualified_range(range_str))
    return frozenset(keys)


def setter_keys_from_bindings(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    workbook: Path | str,
    *,
    export_addresses: Iterable[str] | None = None,
) -> frozenset[NodeKey]:
    """Return sheet-qualified addresses for resolved series-binding setter leaves."""
    from excel_grapher.series_bindings.resolve import resolve_series_bindings

    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="input",
        export_addresses=export_addresses,
    )
    keys: set[NodeKey] = set()
    for result in report.get("series", []):
        for leaf in result.get("leaves", []):
            address = leaf.get("address")
            if isinstance(address, str) and address:
                keys.add(address)
    return frozenset(keys)


def codegen_boundary_keys(
    graph: DependencyGraph,
    *,
    input_ranges: Sequence[str] | None = None,
    series_bindings: WorkbookSeriesBindings | None = None,
    bindings_workbook: Path | str | None = None,
    export_addresses: Iterable[str] | None = None,
) -> frozenset[NodeKey]:
    """Union input leaves, declared ranges, and setter leaves for codegen boundaries."""
    keys: set[NodeKey] = set(input_keys_from_graph(graph))
    keys.update(input_keys_from_ranges(input_ranges))
    if series_bindings is not None:
        if bindings_workbook is None:
            raise ValueError("bindings_workbook is required when series_bindings is set")
        keys.update(
            setter_keys_from_bindings(
                graph,
                series_bindings,
                bindings_workbook,
                export_addresses=export_addresses,
            )
        )
    return frozenset(keys)


def _expand_sheet_qualified_range(range_str: str) -> list[NodeKey]:
    """Expand one sheet-qualified A1 or A1:B2 range to cell keys."""
    if not isinstance(range_str, str):
        raise TypeError("range entries must be strings")
    if "!" not in range_str:
        raise ValueError(f"Range must be sheet-qualified: {range_str}")
    sheet_part, cell_part = range_str.rsplit("!", 1)
    if ":" in cell_part:
        start_cell, end_cell = cell_part.split(":", 1)
    else:
        start_cell = end_cell = cell_part

    sheet, start = parse_address(f"{sheet_part}!{start_cell}")
    _, end = parse_address(f"{sheet_part}!{end_cell}")

    start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start)
    end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end)
    start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
    end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)

    r1, r2 = (start_row, end_row) if start_row <= end_row else (end_row, start_row)
    c1, c2 = (
        (start_col_idx, end_col_idx)
        if start_col_idx <= end_col_idx
        else (end_col_idx, start_col_idx)
    )
    out: list[NodeKey] = []
    for row in range(r1, r2 + 1):
        for col_idx in range(c1, c2 + 1):
            col = fastpyxl.utils.cell.get_column_letter(col_idx)
            out.append(format_cell_key(sheet, col, row))
    return out
