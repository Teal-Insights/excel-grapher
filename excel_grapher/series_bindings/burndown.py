"""Coverage worklist for unbound internal formula cells.

Burndown is a residual report inside the current bound graph closure: after
semantic authoring, it shows which formula nodes already on that graph are
still unbound. It is not a full workbook walk. Empty bindings yield an empty
graph and a zero unbound count. Collapsed A1 rectangles are a coverage
worklist, not candidate bindings to stamp as YAML.
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Literal

from fastpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)

from excel_grapher.core.address_keys import (
    format_cell_key,
    format_range_key,
    parse_address,
    split_address_on_colon,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.occupancy import binding_direction, occupancy_addresses
from excel_grapher.series_bindings.ranges import expand_bound_series_addresses_for_graph
from excel_grapher.series_bindings.resolve import BindingDirection
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

LayoutHint = Literal["scalar", "series", "matrix"]


@dataclass(frozen=True)
class BurndownSheetSummary:
    """Unbound formula-cell counts for one worksheet."""

    sheet: str
    unbound_count: int
    row_count: int

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this summary."""
        return asdict(self)


@dataclass(frozen=True)
class BurndownRow:
    """One row of unbound formula cells in the coverage worklist."""

    sheet: str
    row: int
    spans: str
    layout_hint: LayoutHint

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this row."""
        return asdict(self)


@dataclass(frozen=True)
class BindingBurndownReport:
    """Unbound internal-formula coverage residual for a workbook graph."""

    formula_node_count: int
    unbound_count: int
    unbound_cells: tuple[str, ...]
    collapsed_ranges: tuple[str, ...]
    collapsed_layout_hints: tuple[LayoutHint, ...]
    sheets: tuple[BurndownSheetSummary, ...]
    rows: tuple[BurndownRow, ...]
    rows_truncated: bool = False

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this report."""
        return {
            "formula_node_count": self.formula_node_count,
            "unbound_count": self.unbound_count,
            "unbound_cells": list(self.unbound_cells),
            "collapsed_ranges": list(self.collapsed_ranges),
            "collapsed_layout_hints": list(self.collapsed_layout_hints),
            "sheets": [sheet.to_dict() for sheet in self.sheets],
            "rows": [row.to_dict() for row in self.rows],
            "rows_truncated": self.rows_truncated,
        }


def layout_hint_for_span(*, width: int, height: int) -> LayoutHint:
    """Return a weak layout hint for a rectangular coverage span.

    This is advisory only. Prefer `matrix` when a semantic family has both a
    row-label dimension and a header dimension; do not stamp one series per
    printed row.
    """
    if width <= 1 and height <= 1:
        return "scalar"
    if width > 1 and height > 1:
        return "matrix"
    return "series"


def layout_hint_for_range(data_range: str) -> LayoutHint:
    """Return a weak layout hint for a collapsed A1 coverage rectangle."""
    split = split_address_on_colon(data_range)
    if split is None:
        return "scalar"
    start, end = split
    start_sheet, start_coord = parse_address(start)
    if "!" in end:
        _end_sheet, end_coord = parse_address(end)
    else:
        _, end_coord = parse_address(f"{start_sheet}!{end}")
    start_col, start_row = coordinate_from_string(start_coord)
    end_col, end_row = coordinate_from_string(end_coord)
    width = abs(column_index_from_string(end_col) - column_index_from_string(start_col)) + 1
    height = abs(int(end_row) - int(start_row)) + 1
    return layout_hint_for_span(width=width, height=height)


def contiguous_column_ranges(columns: Sequence[int]) -> list[tuple[int, int]]:
    """Group sorted 1-based column indices into contiguous spans."""
    if not columns:
        return []
    ordered = sorted(columns)
    ranges: list[tuple[int, int]] = []
    start = prev = ordered[0]
    for column in ordered[1:]:
        if column == prev + 1:
            prev = column
            continue
        ranges.append((start, prev))
        start = prev = column
    ranges.append((start, prev))
    return ranges


def group_unbound_cells_by_sheet_row(
    unbound_cells: Sequence[str],
) -> dict[str, dict[int, list[int]]]:
    """Bucket unbound sheet-qualified addresses by sheet and row."""
    grouped: dict[str, dict[int, list[int]]] = defaultdict(lambda: defaultdict(list))
    for address in unbound_cells:
        sheet, coord = parse_address(address)
        column_letters, row = coordinate_from_string(coord)
        grouped[sheet][int(row)].append(column_index_from_string(column_letters))
    return {sheet: dict(rows) for sheet, rows in grouped.items()}


def collapse_unbound_cells_to_ranges(unbound_cells: Sequence[str]) -> tuple[str, ...]:
    """Collapse addresses into maximal sheet-qualified A1 rectangles.

    Consecutive rows that share the same column-span signature become one
    rectangle per span (`Sheet!B2:C10`). Gaps in rows or columns split.
    """
    grouped = group_unbound_cells_by_sheet_row(unbound_cells)
    ranges: list[str] = []
    for sheet, rows in sorted(grouped.items()):
        signatures: list[tuple[int, tuple[tuple[int, int], ...]]] = []
        for row in sorted(rows):
            spans = tuple(contiguous_column_ranges(rows[row]))
            signatures.append((row, spans))
        index = 0
        while index < len(signatures):
            start_row, spans = signatures[index]
            end_row = start_row
            cursor = index + 1
            while (
                cursor < len(signatures)
                and signatures[cursor][0] == end_row + 1
                and signatures[cursor][1] == spans
            ):
                end_row = signatures[cursor][0]
                cursor += 1
            for column_start, column_end in spans:
                start_letter = get_column_letter(column_start)
                end_letter = get_column_letter(column_end)
                start_cell = f"{start_letter}{start_row}"
                end_cell = f"{end_letter}{end_row}"
                if start_cell == end_cell:
                    ranges.append(format_cell_key(sheet, start_letter, start_row))
                else:
                    ranges.append(format_range_key(sheet, start_cell, end_cell))
            index = cursor
    return tuple(ranges)


def format_row_column_spans(
    *,
    sheet: str,
    row: int,
    columns: Sequence[int],
) -> str:
    """Format one row's contiguous column spans as sheet-qualified ranges."""
    spans = ", ".join(
        format_cell_key(sheet, get_column_letter(start), row)
        if start == end
        else format_range_key(
            sheet,
            f"{get_column_letter(start)}{row}",
            f"{get_column_letter(end)}{row}",
        )
        for start, end in contiguous_column_ranges(columns)
    )
    return spans


def required_internal_formula_cells(
    graph: DependencyGraph,
    *,
    input_cells: Iterable[str],
    output_cells: Iterable[str],
) -> tuple[str, ...]:
    """Return formula nodes that are not already covered by public I/O bindings."""
    excluded_cells = set(input_cells) | set(output_cells)
    return tuple(address for address in graph.formula_keys() if address not in excluded_cells)


def manifest_binding_addresses(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    direction: BindingDirection,
    workbook: Path | str | None = None,
) -> set[str]:
    """Expand binding manifests to occupancy addresses for one direction."""
    addresses: set[str] = set()
    for series in bindings.get("series", []):
        if not isinstance(series, dict) or binding_direction(series) != direction:
            continue
        try:
            expanded = expand_bound_series_addresses_for_graph(graph, series, workbook=workbook)
        except (ValueError, TypeError):
            continue
        if not expanded:
            continue
        addresses.update(occupancy_addresses(graph, series, expanded))
    return addresses


def find_unbound_internal_formula_cells(
    *,
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    exempt_cells: frozenset[str] = frozenset(),
    workbook: Path | str | None = None,
) -> tuple[str, ...]:
    """Return formula nodes still missing input, output, or internal coverage."""
    input_cells = manifest_binding_addresses(graph, bindings, direction="input", workbook=workbook)
    output_cells = manifest_binding_addresses(
        graph, bindings, direction="output", workbook=workbook
    )
    internal_cells = manifest_binding_addresses(
        graph, bindings, direction="internal", workbook=workbook
    )
    unbound: list[str] = []
    for address in required_internal_formula_cells(
        graph,
        input_cells=input_cells,
        output_cells=output_cells,
    ):
        if address in exempt_cells or address in internal_cells:
            continue
        unbound.append(address)
    return tuple(unbound)


def load_exempt_addresses(path: Path) -> frozenset[str]:
    """Load reviewed sheet-qualified addresses from a text file.

    Blank lines and `#` comments are ignored.
    """
    addresses: list[str] = []
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        addresses.append(line)
    return frozenset(addresses)


def internal_binding_burndown(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    exempt_cells: frozenset[str] = frozenset(),
    workbook: Path | str | None = None,
    per_sheet: str | None = None,
    max_rows: int | None = None,
) -> BindingBurndownReport:
    """Build the unbound-internal coverage worklist for a graph and sidecar."""
    unbound = find_unbound_internal_formula_cells(
        graph=graph,
        bindings=bindings,
        exempt_cells=exempt_cells,
        workbook=workbook,
    )
    collapsed = collapse_unbound_cells_to_ranges(unbound)
    grouped = group_unbound_cells_by_sheet_row(unbound)
    sheets = tuple(
        BurndownSheetSummary(
            sheet=sheet,
            unbound_count=sum(len(columns) for columns in grouped[sheet].values()),
            row_count=len(grouped[sheet]),
        )
        for sheet in sorted(
            grouped,
            key=lambda name: -sum(len(columns) for columns in grouped[name].values()),
        )
    )
    rows: list[BurndownRow] = []
    rows_truncated = False
    for sheet, sheet_rows in sorted(grouped.items()):
        if per_sheet is not None and sheet != per_sheet:
            continue
        for printed, row in enumerate(sorted(sheet_rows)):
            if max_rows is not None and printed >= max_rows:
                rows_truncated = True
                break
            columns = sheet_rows[row]
            spans = contiguous_column_ranges(columns)
            width = sum(end - start + 1 for start, end in spans)
            rows.append(
                BurndownRow(
                    sheet=sheet,
                    row=row,
                    spans=format_row_column_spans(sheet=sheet, row=row, columns=columns),
                    layout_hint=layout_hint_for_span(width=width, height=1),
                )
            )
    return BindingBurndownReport(
        formula_node_count=len(graph.formula_keys()),
        unbound_count=len(unbound),
        unbound_cells=unbound,
        collapsed_ranges=collapsed,
        collapsed_layout_hints=tuple(layout_hint_for_range(item) for item in collapsed),
        sheets=sheets,
        rows=tuple(rows),
        rows_truncated=rows_truncated,
    )


def format_burndown_report(report: BindingBurndownReport) -> list[str]:
    """Render a coverage worklist as stable, human-readable lines."""
    lines = [
        "Internal-binding coverage worklist (not a generator)",
        "Scope: residual formula nodes in the current bound graph closure, "
        "not a full workbook walk",
        f"Formula nodes: {report.formula_node_count}",
        f"Unbound internal formula cells: {report.unbound_count}",
        f"Collapsed A1 ranges: {len(report.collapsed_ranges)}",
    ]
    for data_range, hint in zip(
        report.collapsed_ranges, report.collapsed_layout_hints, strict=True
    ):
        lines.append(f"  {data_range}  [hint: {hint}]")
    if report.sheets:
        lines.append("")
        lines.append("Unbound cells by sheet:")
        for sheet in report.sheets:
            lines.append(
                f"  {sheet.sheet}: {sheet.unbound_count} cells across {sheet.row_count} rows"
            )
    current_sheet: str | None = None
    for row in report.rows:
        if row.sheet != current_sheet:
            lines.append("")
            lines.append(f"== {row.sheet} ==")
            current_sheet = row.sheet
        lines.append(f"  row {row.row}: {row.spans}  [hint: {row.layout_hint}]")
    if report.rows_truncated:
        lines.append("  ... (truncated)")
    return lines
