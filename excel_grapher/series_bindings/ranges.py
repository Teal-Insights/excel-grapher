"""Range expansion for series binding `data_range` fields."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING

import fastpyxl

from excel_grapher.grapher.parser import format_key
from excel_grapher.grapher.resolver import build_named_range_map
from excel_grapher.grapher.target_expansion import expand_targets_to_roots

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph


def _resolve_named_range_maps(
    *,
    workbook: Path | str | None,
    named_ranges: Mapping[str, tuple[str, str]] | None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None,
) -> tuple[dict[str, tuple[str, str]], dict[str, tuple[str, str, str]], list[str]]:
    nr = dict(named_ranges or {})
    nrr = dict(named_range_ranges or {})
    sheetnames: list[str] = []
    if workbook is not None and (not nr and not nrr):
        path = Path(workbook)
        keep_vba = path.suffix.lower() == ".xlsm"
        wb = fastpyxl.load_workbook(path, data_only=False, read_only=True, keep_vba=keep_vba)
        sheetnames = list(wb.sheetnames)
        maps = build_named_range_map(wb)
        nr = dict(maps.cell_map)
        nrr = dict(maps.range_map)
    return nr, nrr, sheetnames


def _sheetnames_for_target(
    data_range: str,
    *,
    sheetnames: Sequence[str] | None,
    named_ranges: Mapping[str, tuple[str, str]],
    named_range_ranges: Mapping[str, tuple[str, str, str]],
) -> list[str]:
    if sheetnames:
        return list(sheetnames)
    if "!" in data_range:
        from excel_grapher.core.address_keys import parse_address
        from excel_grapher.grapher.target_expansion import split_range_target_on_colon

        split = split_range_target_on_colon(data_range)
        start = split[0] if split is not None else data_range
        sheet, _ = parse_address(start)
        return [sheet]
    if data_range in named_range_ranges:
        return [named_range_ranges[data_range][0]]
    if data_range in named_ranges:
        return [named_ranges[data_range][0]]
    raise ValueError(
        f"Cannot infer sheet for data_range {data_range!r}; pass sheetnames= or workbook= "
        "for defined-name targets"
    )


def expand_data_range(
    data_range: str,
    *,
    workbook: Path | str | None = None,
    sheetnames: Sequence[str] | None = None,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
    max_range_cells: int = 5000,
) -> list[str]:
    """Expand a binding `data_range` to canonical sheet-qualified cell addresses.

    Uses the same target expansion as `create_dependency_graph` (including
    both-end sheet-qualified ranges like `Sheet1!A1:Sheet1!B2` and defined names).
    """
    nr, nrr, wb_sheets = _resolve_named_range_maps(
        workbook=workbook,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )
    if "!" not in data_range and data_range not in nr and data_range not in nrr:
        raise ValueError(
            f"Unknown defined name {data_range!r}; pass workbook= or named-range maps from the graph"
        )
    sheets = _sheetnames_for_target(
        data_range,
        sheetnames=sheetnames or (wb_sheets if wb_sheets else None),
        named_ranges=nr,
        named_range_ranges=nrr,
    )

    roots = expand_targets_to_roots(
        [data_range],
        sheetnames=sheets,
        named_ranges=nr,
        named_range_ranges=nrr,
        max_range_cells=max_range_cells,
    )
    return [format_key(sheet, a1) for sheet, a1 in roots]


def expand_data_range_for_graph(
    graph: DependencyGraph,
    data_range: str,
    *,
    workbook: Path | str | None = None,
    max_range_cells: int = 5000,
) -> list[str]:
    """Expand `data_range` using named-range maps (and optional workbook) from a graph."""
    return expand_data_range(
        data_range,
        workbook=workbook,
        sheetnames=graph.sheet_order,
        named_ranges=graph.named_ranges,
        named_range_ranges=graph.named_range_ranges,
        max_range_cells=max_range_cells,
    )
