from __future__ import annotations

from datetime import datetime
from typing import Any, Literal, NotRequired, TypedDict

ValidationLevel = Literal["error", "warning"]
Scalar = str | int | float | bool | datetime | None
Record = dict[str, object]
Records = list[Record]


class WorkbookSeriesBindings(TypedDict):
    """Top-level workbook binding manifest (post-merge, post-schema validation)."""

    schema_version: str
    series: list[dict[str, Any]]
    workbook: NotRequired[str]
    concept_scheme: NotRequired[dict[str, Any]]


class ValidationIssue(TypedDict):
    level: ValidationLevel
    code: str
    message: str
    series_id: str | None
    address: str | None


def make_issue(
    level: Literal["error", "warning"],
    code: str,
    message: str,
    *,
    series_id: str | None = None,
    address: str | None = None,
) -> ValidationIssue:
    return {
        "level": level,
        "code": code,
        "message": message,
        "series_id": series_id,
        "address": address,
    }


class ValidationReport(TypedDict):
    ok: bool
    issues: list[ValidationIssue]


class LeafResolution(TypedDict):
    address: str
    coordinates: dict[str, Scalar]
    key: dict[str, Scalar]
    record: dict[str, Scalar]


class ResolutionIssue(ValidationIssue):
    """Alias for resolution-time validation issues."""


class SeriesResolution(TypedDict):
    series_id: str
    ok: bool
    requires_address: bool
    leaves: list[LeafResolution]
    issues: list[ResolutionIssue]


class ResolutionReport(TypedDict):
    ok: bool
    series: list[SeriesResolution]
    issues: list[ResolutionIssue]


class InputSeriesCell(TypedDict):
    address: str
    coordinates: dict[str, Scalar]
    key: dict[str, Scalar]
    record: dict[str, Scalar]


class InputSeries(TypedDict):
    id: str
    setter_name: str
    key_fields: list[str]
    requires_address: bool
    cells: list[InputSeriesCell]
    issues: list[ResolutionIssue]


class OutputSeriesCell(TypedDict):
    address: str
    coordinates: dict[str, Scalar]
    key: dict[str, Scalar]
    record: dict[str, Scalar]


class OutputSeries(TypedDict):
    id: str
    compute_name: str
    key_fields: list[str]
    cells: list[OutputSeriesCell]
    issues: list[ResolutionIssue]


class InternalSeriesCell(TypedDict):
    address: str
    coordinates: dict[str, Scalar]
    key: dict[str, Scalar]
    record: dict[str, Scalar]


class InternalSeries(TypedDict):
    id: str
    key_fields: list[str]
    requires_address: bool
    cells: list[InternalSeriesCell]
    issues: list[ResolutionIssue]


class ConstantSeriesCell(TypedDict):
    address: str
    coordinates: dict[str, Scalar]
    key: dict[str, Scalar]
    record: dict[str, Scalar]


class ConstantSeries(TypedDict):
    id: str
    reader_name: str
    key_fields: list[str]
    requires_address: bool
    cells: list[ConstantSeriesCell]
    issues: list[ResolutionIssue]
