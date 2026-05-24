"""Input-group discovery from dependency graph leaf inputs."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Mapping, Sequence
from typing import TYPE_CHECKING

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    normalize_key,
    normalize_range_key,
    quote_sheet_if_needed,
)
from excel_grapher.exporter.input_groups import (
    BoundingBox,
    GroupingOptions,
    InputCell,
    InputGroup,
    InputGroupsPayload,
    InputGroupsSummary,
    LabelMode,
    Orientation,
    address_in_range,
    project_labels,
    stable_group_id,
    utc_now_iso,
)

if TYPE_CHECKING:
    from excel_grapher.exporter.codegen import CodeGenerator

_DEFAULT_ORIENTATION: Orientation = "rowwise"
_ORIENTATION_RANK = {"rowwise": 0, "columnwise": 1}


def _label_key(labels: Sequence[str]) -> tuple[str, ...]:
    return tuple(labels)


def _resolve_orientation(
    address: str,
    overrides: Sequence[tuple[str, Orientation]],
) -> Orientation:
    orientation = _DEFAULT_ORIENTATION
    for range_spec, override_orientation in overrides:
        if address_in_range(address, range_spec):
            orientation = override_orientation
    return orientation


def _build_override_pairs(options: GroupingOptions) -> list[tuple[str, Orientation]]:
    return [(o.range_spec, o.orientation) for o in options.overrides]


def _override_range_for_address(
    address: str,
    overrides: Sequence[tuple[str, Orientation]],
) -> str | None:
    matched: str | None = None
    for range_spec, _orientation in overrides:
        if address_in_range(address, range_spec):
            matched = range_spec
    return matched


def _compute_bounding_box(cells: Sequence[InputCell]) -> BoundingBox:
    rows = [c.row for c in cells]
    cols = [c.col for c in cells]
    return BoundingBox(min(rows), min(cols), max(rows), max(cols))


def _is_full_rectangle(cells: Sequence[InputCell], bbox: BoundingBox) -> bool:
    expected = bbox.row_count * bbox.col_count
    if len(cells) != expected:
        return False
    positions = {(c.row, c.col) for c in cells}
    for row in range(bbox.min_row, bbox.max_row + 1):
        for col in range(bbox.min_col, bbox.max_col + 1):
            if (row, col) not in positions:
                return False
    return True


def _format_range_a1(sheet: str, bbox: BoundingBox) -> str:
    start_col = fastpyxl.utils.cell.get_column_letter(bbox.min_col)
    end_col = fastpyxl.utils.cell.get_column_letter(bbox.max_col)
    sheet_q = quote_sheet_if_needed(sheet)
    start = f"{start_col}{bbox.min_row}"
    end = f"{end_col}{bbox.max_row}"
    if start == end:
        return f"{sheet_q}!{start}"
    return normalize_range_key(f"{sheet_q}!{start}:{end}")


def _grouping_key(
    cell: InputCell,
    orientation: Orientation,
    override_range: str | None,
) -> tuple[str, Orientation, tuple[str, ...], tuple[str, ...]]:
    row_key = _label_key(cell.row_labels)
    col_key = _label_key(cell.column_labels)
    if row_key:
        return (cell.sheet, orientation, row_key, ())
    if col_key:
        return (cell.sheet, orientation, (), col_key)
    if override_range is not None:
        return (cell.sheet, orientation, (override_range,), ())
    return (cell.sheet, orientation, (cell.address,), ())


def _read_metadata_labels(
    metadata: Mapping[str, object] | None,
    *,
    include_labels: bool,
    label_mode: LabelMode,
) -> tuple[tuple[str, ...], tuple[str, ...]]:
    if not include_labels or metadata is None:
        return (), ()
    raw_row = metadata.get("row_labels", ())
    raw_col = metadata.get("column_labels", ())
    row_labels = tuple(str(x) for x in raw_row) if isinstance(raw_row, (list, tuple)) else ()
    col_labels = tuple(str(x) for x in raw_col) if isinstance(raw_col, (list, tuple)) else ()
    return (
        project_labels(row_labels, label_mode),
        project_labels(col_labels, label_mode),
    )


def discover_input_groups_from_graph(
    generator: CodeGenerator,
    targets: Sequence[str],
    *,
    grouping: GroupingOptions | None = None,
    constant_types: set[str] | None = None,
    constant_ranges: Sequence[str] | None = None,
    constant_blanks: bool = False,
    input_ranges: Sequence[str] | None = None,
    workbook_name: str | None = None,
) -> InputGroupsPayload:
    options = grouping or GroupingOptions()
    label_mode = options.effective_label_mode()
    override_pairs = _build_override_pairs(options)

    normalized_targets = [normalize_key(t) for t in targets]
    inputs, _constants = generator.classify_leaf_nodes(
        list(normalized_targets),
        constant_types=constant_types,
        constant_ranges=constant_ranges,
        constant_blanks=constant_blanks,
        input_ranges=input_ranges,
    )

    cells: list[tuple[InputCell, Orientation, tuple[str, ...], tuple[str, ...], str | None]] = []
    for address in sorted(inputs, key=lambda a: parse_address_key(a)):
        node = generator.graph.get_node(address)
        metadata = None if node is None else getattr(node, "metadata", None)
        row_labels, col_labels = _read_metadata_labels(
            metadata,
            include_labels=options.include_labels,
            label_mode=label_mode,
        )
        cell = InputCell.from_address(
            address,
            row_labels=row_labels,
            column_labels=col_labels,
        )
        orientation = _resolve_orientation(address, override_pairs)
        override_range = _override_range_for_address(address, override_pairs)
        row_key = _label_key(cell.row_labels)
        col_key = _label_key(cell.column_labels)
        cells.append((cell, orientation, row_key, col_key, override_range))

    partitions: dict[
        tuple[str, Orientation, tuple[str, ...], tuple[str, ...]],
        list[InputCell],
    ] = defaultdict(list)
    for cell, orientation, _row_key, _col_key, override_range in cells:
        key = _grouping_key(cell, orientation, override_range)
        partitions[key].append(cell)

    sheet_order = _graph_sheet_order(generator)
    groups: list[InputGroup] = []
    for key in sorted(
        partitions.keys(),
        key=lambda k: (
            sheet_order.get(k[0], len(sheet_order)),
            k[0],
            _ORIENTATION_RANK[k[1]],
            k[2],
            k[3],
        ),
    ):
        sheet, orientation, row_labels_key, column_labels_key = key
        group_cells = sorted(partitions[key], key=lambda c: (c.row, c.col))
        bbox = _compute_bounding_box(group_cells)
        rectangular = _is_full_rectangle(group_cells, bbox)
        group = InputGroup(
            group_id=stable_group_id(sheet, orientation, row_labels_key, column_labels_key),
            sheet=sheet,
            orientation=orientation,
            row_labels_key=row_labels_key,
            column_labels_key=column_labels_key,
            cells=tuple(group_cells),
            bounding_box=bbox if rectangular else None,
            shape=(bbox.row_count, bbox.col_count) if rectangular else None,
            range_a1=_format_range_a1(sheet, bbox) if rectangular else None,
        )
        groups.append(group)

    hist: dict[Orientation, int] = {"rowwise": 0, "columnwise": 0}
    total_cells = 0
    for group in groups:
        hist[group.orientation] = hist.get(group.orientation, 0) + 1
        total_cells += len(group.cells)

    return InputGroupsPayload(
        workbook_name=workbook_name,
        generated_at_utc=utc_now_iso(),
        summary=InputGroupsSummary(
            total_groups=len(groups),
            total_cells=total_cells,
            orientation_histogram=hist,
        ),
        groups=tuple(groups),
    )


def parse_address_key(address: str) -> tuple[int, str, int, int]:
    from excel_grapher.core.address_keys import parse_address

    sheet, cell = parse_address(normalize_key(address))
    col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
    col = fastpyxl.utils.cell.column_index_from_string(col_str)
    return (0, sheet, int(row), col)


def _graph_sheet_order(generator: CodeGenerator) -> dict[str, int]:
    sheetnames = getattr(generator.graph, "sheetnames", None)
    if not isinstance(sheetnames, list):
        return {}
    return {str(name): idx for idx, name in enumerate(sheetnames)}


def validate_input_groups(groups: Sequence[InputGroup]) -> None:
    seen_ids: set[str] = set()
    seen_addresses: set[str] = set()
    for group in groups:
        if group.group_id in seen_ids:
            raise ValueError(f"Duplicate group_id: {group.group_id!r}")
        seen_ids.add(group.group_id)
        for cell in group.cells:
            if cell.address in seen_addresses:
                raise ValueError(f"Duplicate cell address across groups: {cell.address!r}")
            seen_addresses.add(cell.address)
        if group.bounding_box is not None and group.shape is not None:
            bbox = group.bounding_box
            if not _is_full_rectangle(group.cells, bbox):
                raise ValueError(f"Group {group.group_id!r} bounding_box does not match cells")
