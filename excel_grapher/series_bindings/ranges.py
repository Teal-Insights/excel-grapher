"""Range expansion for series binding `data_range` fields."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any

import fastpyxl
from fastpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_grapher.core.address_keys import format_range_key, parse_address
from excel_grapher.grapher.parser import format_key
from excel_grapher.grapher.resolver import build_named_range_map
from excel_grapher.grapher.target_expansion import expand_targets_to_roots
from excel_grapher.series_bindings.geometry import expand_column_specs, expand_row_specs

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
    both-end sheet-qualified ranges like `Sheet1!A1:Sheet1!B2`, which collapse
    to single-prefix form, and defined names).
    """
    # Sheet-qualified targets do not need named-range maps; skip the workbook open.
    nr, nrr, wb_sheets = _resolve_named_range_maps(
        workbook=None if "!" in data_range else workbook,
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


def _parse_cell_rc(address: str) -> tuple[str, int, int]:
    sheet, coord = parse_address(address)
    col_letters, row = fastpyxl.utils.cell.coordinate_from_string(coord)
    return sheet, int(row), column_index_from_string(col_letters)


def _solid_rectangle_address(addresses: Sequence[str]) -> str | None:
    """Return a sheet-qualified A1 range when `addresses` form one solid rectangle."""
    if not addresses:
        return None
    cells = [_parse_cell_rc(address) for address in addresses]
    sheets = {sheet for sheet, _row, _col in cells}
    if len(sheets) != 1:
        return None
    sheet = next(iter(sheets))
    rows = {row for _sheet, row, _col in cells}
    cols = {col for _sheet, _row, col in cells}
    min_row, max_row = min(rows), max(rows)
    min_col, max_col = min(cols), max(cols)
    expected = (max_row - min_row + 1) * (max_col - min_col + 1)
    if len(cells) != expected:
        return None
    cell_set = {(row, col) for _sheet, row, col in cells}
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if (row, col) not in cell_set:
                return None
    start = f"{get_column_letter(min_col)}{min_row}"
    end = f"{get_column_letter(max_col)}{max_row}"
    return format_range_key(sheet, start, end)


def effective_reader_range_address(
    series: Mapping[str, Any],
    *,
    workbook: Path | str | None = None,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
    max_range_cells: int = 5000,
) -> str | None:
    """Return the address for `read_*_range`, or None when none should be emitted.

    Without `exclude_rows` / `exclude_columns`, returns the series `data_range`.
    With exclusions, expands `data_range`, drops excluded rows/columns, and returns
    a single contiguous rectangle covering exactly the remaining cells — or None
    when the selection is empty or cannot be expressed as one `xl_range`.
    """
    data_range = series.get("data_range")
    if not isinstance(data_range, str) or not data_range:
        return None
    exclude_rows = series.get("exclude_rows")
    exclude_columns = series.get("exclude_columns")
    if not exclude_rows and not exclude_columns:
        return data_range

    addresses = expand_data_range(
        data_range,
        workbook=workbook,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
        max_range_cells=max_range_cells,
    )
    if exclude_rows:
        excluded_rows = expand_row_specs(exclude_rows)
        addresses = [
            address for address in addresses if _parse_cell_rc(address)[1] not in excluded_rows
        ]
    if exclude_columns:
        excluded_cols = expand_column_specs(exclude_columns)
        addresses = [
            address for address in addresses if _parse_cell_rc(address)[2] not in excluded_cols
        ]
    return _solid_rectangle_address(addresses)
