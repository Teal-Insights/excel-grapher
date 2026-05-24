"""Typed contracts for input-group discovery and setter code generation."""

from __future__ import annotations

import hashlib
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any, Literal, TypeAlias, cast

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_cell_key, normalize_key, parse_address
from excel_grapher.grapher.blank_ranges import parse_blank_range_spec

Scalar: TypeAlias = int | float | str | bool | None
Record: TypeAlias = dict[str, Scalar | list[str]]
Records: TypeAlias = list[Record]

LabelMode = Literal["none", "first", "all"]
Orientation = Literal["rowwise", "columnwise"]
TargetShape = Literal["cell", "row_vector", "col_vector", "rectangle"]


def validate_record(record: Mapping[str, object], *, strict: bool = True) -> None:
    if "value" not in record:
        raise ValueError("Record must include 'value'")
    if not strict:
        return
    for key, val in record.items():
        if isinstance(val, list):
            if not all(isinstance(item, str) for item in val):
                raise ValueError(f"Record field {key!r} must be a list of strings")
        elif val is not None and not isinstance(val, (int, float, str, bool)):
            raise ValueError(f"Record field {key!r} has unsupported scalar type")


def validate_records(records: Sequence[Mapping[str, object]], *, strict: bool = True) -> None:
    for record in records:
        validate_record(record, strict=strict)


def _parse_range_bounds(range_spec: str) -> tuple[str, int, int, int, int]:
    sheet, r1, c1, r2, c2 = parse_blank_range_spec(range_spec)
    return sheet, r1, c1, r2, c2


def address_in_range(address: str, range_spec: str) -> bool:
    normalized = normalize_key(address)
    sheet, cell = parse_address(normalized)
    col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
    col = fastpyxl.utils.cell.column_index_from_string(col_str)
    range_sheet, r1, c1, r2, c2 = _parse_range_bounds(range_spec)
    if sheet != range_sheet:
        return False
    return r1 <= row <= r2 and c1 <= col <= c2


@dataclass(frozen=True, slots=True)
class InputCell:
    address: str
    sheet: str
    row: int
    col: int
    row_labels: tuple[str, ...] = ()
    column_labels: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        normalized = normalize_key(self.address)
        if normalized != self.address:
            raise ValueError(f"InputCell address must be normalized: {self.address!r}")
        parsed_sheet, cell = parse_address(normalized)
        col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
        col = fastpyxl.utils.cell.column_index_from_string(col_str)
        if parsed_sheet != self.sheet or int(row) != self.row or col != self.col:
            raise ValueError(
                f"InputCell component mismatch for {self.address!r}: "
                f"expected ({parsed_sheet!r}, {row}, {col}), "
                f"got ({self.sheet!r}, {self.row}, {self.col})"
            )

    @classmethod
    def from_address(
        cls,
        address: str,
        *,
        row_labels: Sequence[str] = (),
        column_labels: Sequence[str] = (),
    ) -> InputCell:
        normalized = normalize_key(address)
        sheet, cell = parse_address(normalized)
        col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
        col = fastpyxl.utils.cell.column_index_from_string(col_str)
        return cls(
            address=normalized,
            sheet=sheet,
            row=int(row),
            col=col,
            row_labels=tuple(row_labels),
            column_labels=tuple(column_labels),
        )


@dataclass(frozen=True, slots=True)
class BoundingBox:
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def row_count(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def col_count(self) -> int:
        return self.max_col - self.min_col + 1

    def to_dict(self) -> dict[str, int]:
        return {
            "min_row": self.min_row,
            "min_col": self.min_col,
            "max_row": self.max_row,
            "max_col": self.max_col,
        }


@dataclass(frozen=True, slots=True)
class InputGroup:
    group_id: str
    sheet: str
    orientation: Orientation
    row_labels_key: tuple[str, ...]
    column_labels_key: tuple[str, ...]
    cells: tuple[InputCell, ...]
    bounding_box: BoundingBox | None = None
    shape: tuple[int, int] | None = None
    range_a1: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "group_id": self.group_id,
            "sheet": self.sheet,
            "orientation": self.orientation,
            "row_labels_key": list(self.row_labels_key),
            "column_labels_key": list(self.column_labels_key),
            "cells": [
                {
                    "address": c.address,
                    "sheet": c.sheet,
                    "row": c.row,
                    "col": c.col,
                    "row_labels": list(c.row_labels),
                    "column_labels": list(c.column_labels),
                }
                for c in self.cells
            ],
            "bounding_box": None
            if self.bounding_box is None
            else {
                "min_row": self.bounding_box.min_row,
                "min_col": self.bounding_box.min_col,
                "max_row": self.bounding_box.max_row,
                "max_col": self.bounding_box.max_col,
            },
            "shape": None if self.shape is None else list(self.shape),
            "range_a1": self.range_a1,
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> InputGroup:
        bbox_raw = data.get("bounding_box")
        bbox = None
        if bbox_raw is not None:
            bbox = BoundingBox(
                min_row=int(bbox_raw["min_row"]),
                min_col=int(bbox_raw["min_col"]),
                max_row=int(bbox_raw["max_row"]),
                max_col=int(bbox_raw["max_col"]),
            )
        shape_raw = data.get("shape")
        shape = None if shape_raw is None else (int(shape_raw[0]), int(shape_raw[1]))
        cells = tuple(
            InputCell(
                address=str(c["address"]),
                sheet=str(c["sheet"]),
                row=int(c["row"]),
                col=int(c["col"]),
                row_labels=tuple(c.get("row_labels", ())),
                column_labels=tuple(c.get("column_labels", ())),
            )
            for c in data["cells"]
        )
        return cls(
            group_id=str(data["group_id"]),
            sheet=str(data["sheet"]),
            orientation=cast("Orientation", str(data["orientation"])),
            row_labels_key=tuple(data.get("row_labels_key", ())),
            column_labels_key=tuple(data.get("column_labels_key", ())),
            cells=cells,
            bounding_box=bbox,
            shape=shape,
            range_a1=data.get("range_a1"),
        )


@dataclass(frozen=True, slots=True)
class InputGroupsSummary:
    total_groups: int
    total_cells: int
    orientation_histogram: dict[Orientation, int]

    def to_dict(self) -> dict[str, Any]:
        return {
            "total_groups": self.total_groups,
            "total_cells": self.total_cells,
            "orientation_histogram": dict(self.orientation_histogram),
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> InputGroupsSummary:
        hist = data.get("orientation_histogram", {})
        orientation_histogram: dict[Orientation, int] = {
            cast("Orientation", str(k)): int(v) for k, v in hist.items()
        }
        return cls(
            total_groups=int(data["total_groups"]),
            total_cells=int(data["total_cells"]),
            orientation_histogram=orientation_histogram,
        )


@dataclass(frozen=True, slots=True)
class InputGroupsPayload:
    workbook_name: str | None
    generated_at_utc: str | None
    summary: InputGroupsSummary
    groups: tuple[InputGroup, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "workbook_name": self.workbook_name,
            "generated_at_utc": self.generated_at_utc,
            "summary": self.summary.to_dict(),
            "groups": [g.to_dict() for g in self.groups],
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> InputGroupsPayload:
        return cls(
            workbook_name=data.get("workbook_name"),
            generated_at_utc=data.get("generated_at_utc"),
            summary=InputGroupsSummary.from_dict(data["summary"]),
            groups=tuple(InputGroup.from_dict(g) for g in data["groups"]),
        )


@dataclass(frozen=True, slots=True)
class GroupingOverride:
    range_spec: str
    orientation: Orientation
    label_mode: LabelMode | None = None
    group_name: str | None = None

    def __post_init__(self) -> None:
        _parse_range_bounds(self.range_spec)


@dataclass(frozen=True, slots=True)
class GroupingOptions:
    include_labels: bool = False
    label_mode: LabelMode = "none"
    overrides: tuple[GroupingOverride, ...] = ()

    def effective_label_mode(self) -> LabelMode:
        if not self.include_labels:
            return "none"
        return self.label_mode


@dataclass(frozen=True, slots=True)
class SetterGenerationOptions:
    include_labels: bool = False
    label_mode: LabelMode = "none"
    include_address: bool = True
    include_position_fields: bool = False
    naming_strategy: Callable[[InputGroup], str] | None = None
    strict_records: bool = True
    grouping: GroupingOptions | None = None

    def effective_label_mode(self) -> LabelMode:
        if not self.include_labels:
            return "none"
        return self.label_mode

    def resolved_grouping(self) -> GroupingOptions:
        if self.grouping is not None:
            return self.grouping
        return GroupingOptions(
            include_labels=self.include_labels,
            label_mode=self.label_mode,
        )


@dataclass(frozen=True, slots=True)
class NormalizedTargetSpec:
    address_or_range: str
    shape: TargetShape
    sheet: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def row_count(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def col_count(self) -> int:
        return self.max_col - self.min_col + 1

    def cell_addresses_row_major(self) -> tuple[str, ...]:
        out: list[str] = []
        for row in range(self.min_row, self.max_row + 1):
            for col in range(self.min_col, self.max_col + 1):
                col_letter = fastpyxl.utils.cell.get_column_letter(col)
                out.append(format_cell_key(self.sheet, col_letter, row))
        return tuple(out)


def project_labels(labels: Sequence[str], mode: LabelMode) -> tuple[str, ...]:
    if mode == "none" or not labels:
        return ()
    if mode == "first":
        return (labels[0],)
    return tuple(labels)


def stable_group_id(
    sheet: str,
    orientation: Orientation,
    row_labels_key: tuple[str, ...],
    column_labels_key: tuple[str, ...],
) -> str:
    material = "|".join(
        [
            sheet,
            orientation,
            ",".join(row_labels_key),
            ",".join(column_labels_key),
        ]
    )
    digest = hashlib.sha256(material.encode()).hexdigest()[:8]
    slug = re.sub(r"[^a-z0-9]+", "_", material.lower()).strip("_")[:48]
    return f"{slug}_{digest}" if slug else digest


def utc_now_iso() -> str:
    return datetime.now(tz=UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
