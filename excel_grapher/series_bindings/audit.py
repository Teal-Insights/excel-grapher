"""Audit series bindings for resolve failures that codegen would treat as fatal.

`validate_series_bindings` and `derive_*_series` do not require resolution
`ok=True`. Codegen does. This module runs the same resolver and reports
Tier-1 authoring failures (bind errors, empty public series, sparse headers
without `fill`, duplicate formula-cell ownership) before export.
"""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Literal

import fastpyxl
from fastpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)

from excel_grapher.core.address_keys import parse_address
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import (
    has_input_direction,
    has_output_direction,
)
from excel_grapher.series_bindings.occupancy import binding_direction
from excel_grapher.series_bindings.ranges import (
    apply_series_excludes,
    expand_data_range_for_graph,
    series_data_ranges,
)
from excel_grapher.series_bindings.resolve import (
    BindingDirection,
    resolve_series_bindings,
)
from excel_grapher.series_bindings.types import (
    SeriesResolution,
    WorkbookSeriesBindings,
)

AuditSeverity = Literal["error", "warning"]
DIRECTIONS: tuple[BindingDirection, ...] = (
    "input",
    "output",
    "internal",
    "constant",
)
_FORMULA_DIRECTIONS = frozenset({"internal", "output"})


@dataclass(frozen=True)
class AuditFinding:
    """One codegen-fatal or advisory binding-resolution finding."""

    severity: AuditSeverity
    code: str
    series_id: str
    direction: BindingDirection
    message: str
    address: str | None = None

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this finding."""
        return asdict(self)


@dataclass(frozen=True)
class BindingResolutionAuditReport:
    """Tier-1 resolution audit across one or more binding directions."""

    findings: tuple[AuditFinding, ...]

    @property
    def ok(self) -> bool:
        """True when no error-severity findings are present."""
        return not any(finding.severity == "error" for finding in self.findings)

    @property
    def error_count(self) -> int:
        """Number of error-severity findings."""
        return sum(1 for finding in self.findings if finding.severity == "error")

    @property
    def warning_count(self) -> int:
        """Number of warning-severity findings."""
        return sum(1 for finding in self.findings if finding.severity == "warning")

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable mapping of this report."""
        return {
            "ok": self.ok,
            "error_count": self.error_count,
            "warning_count": self.warning_count,
            "findings": [finding.to_dict() for finding in self.findings],
        }


def _series_by_id(bindings: WorkbookSeriesBindings) -> dict[str, dict[str, Any]]:
    return {
        str(series["id"]): series
        for series in bindings.get("series", [])
        if isinstance(series, dict) and "id" in series
    }


def _is_public_api_series(series: dict[str, Any], direction: BindingDirection) -> bool:
    if direction == "input":
        return has_input_direction(series)
    if direction == "output":
        return has_output_direction(series)
    return False


def _is_blank_label(raw: object) -> bool:
    return raw is None or (isinstance(raw, str) and not raw.strip())


def unfilled_label_binds(series: dict[str, Any]) -> list[dict[str, Any]]:
    """Return `column_header` / `row_label` binds that lack `fill: true`."""
    structure = series.get("structure") or {}
    dimensions = structure.get("dimensions") or []
    if not isinstance(dimensions, list):
        return []
    binds: list[dict[str, Any]] = []
    for dim in dimensions:
        if not isinstance(dim, dict):
            continue
        bind = dim.get("bind")
        if not isinstance(bind, dict):
            continue
        if bind.get("kind") not in {"column_header", "row_label"}:
            continue
        if bool(bind.get("fill", False)):
            continue
        binds.append(bind)
    return binds


def _expand_series_addresses_for_graph(
    graph: DependencyGraph,
    series: dict[str, Any],
    *,
    workbook_path: Path | str,
) -> list[str] | None:
    try:
        addresses: list[str] = []
        for data_range in series_data_ranges(series):
            addresses.extend(expand_data_range_for_graph(graph, data_range, workbook=workbook_path))
        return apply_series_excludes(addresses, series)
    except (ValueError, TypeError):
        return None


def _data_range_row_column_indices(
    graph: DependencyGraph,
    series: dict[str, Any],
    *,
    workbook_path: Path | str,
) -> tuple[str, frozenset[int], frozenset[int]] | None:
    addresses = _expand_series_addresses_for_graph(graph, series, workbook_path=workbook_path)
    if not addresses:
        return None

    rows: set[int] = set()
    columns: set[int] = set()
    sheet_name: str | None = None
    for address in addresses:
        sheet, coord = parse_address(address)
        if sheet_name is None:
            sheet_name = sheet
        column_letters, row = coordinate_from_string(coord)
        rows.add(int(row))
        columns.add(column_index_from_string(column_letters))
    if sheet_name is None:
        return None
    return sheet_name, frozenset(rows), frozenset(columns)


def find_sparse_label_bind_issues(
    graph: DependencyGraph,
    series: dict[str, Any],
    *,
    workbook_path: Path | str,
    direction: BindingDirection,
    workbook: fastpyxl.Workbook | None = None,
) -> list[AuditFinding]:
    """Flag `column_header` / `row_label` binds that span blank labels without fill.

    Pass a shared `workbook` from `audit_binding_resolutions` so the xlsx is
    opened once. When `workbook` is omitted this helper loads and closes it.
    """
    binds = unfilled_label_binds(series)
    if not binds:
        return []

    series_id = str(series.get("id", ""))
    span = _data_range_row_column_indices(graph, series, workbook_path=workbook_path)
    if span is None:
        return []
    sheet_name, rows, columns = span

    owns_workbook = workbook is None
    loaded = (
        workbook
        if workbook is not None
        else fastpyxl.load_workbook(Path(workbook_path), data_only=False, read_only=True)
    )
    try:
        if sheet_name not in loaded.sheetnames:
            return []
        worksheet = loaded[sheet_name]
        findings: list[AuditFinding] = []
        for bind in binds:
            kind = bind.get("kind")
            blank_sources: list[str] = []
            if kind == "column_header":
                header_row = int(bind["header_row"])
                for column in sorted(columns):
                    raw = worksheet.cell(header_row, column).value
                    if _is_blank_label(raw):
                        blank_sources.append(
                            f"{sheet_name}!{get_column_letter(column)}{header_row}"
                        )
                axis_label = f"column_header row {header_row}"
            else:
                label_column = str(bind["label_column"])
                label_col_index = column_index_from_string(label_column)
                for row in sorted(rows):
                    raw = worksheet.cell(row, label_col_index).value
                    if _is_blank_label(raw):
                        blank_sources.append(f"{sheet_name}!{label_column}{row}")
                axis_label = f"row_label column {label_column}"

            if not blank_sources:
                continue
            sample = blank_sources[0]
            if len(blank_sources) == 1:
                detail = sample
            else:
                detail = f"{len(blank_sources)} cells ({sample}-{blank_sources[-1]})"
            findings.append(
                AuditFinding(
                    severity="error",
                    code="sparse_label_without_fill",
                    series_id=series_id,
                    direction=direction,
                    message=(
                        f"{axis_label}: blank label(s) inside data_range without "
                        f"fill: true ({detail})"
                    ),
                    address=sample,
                )
            )
        return findings
    finally:
        if owns_workbook:
            loaded.close()


def findings_from_resolution(
    resolved: SeriesResolution,
    *,
    direction: BindingDirection,
    series: dict[str, Any] | None,
) -> list[AuditFinding]:
    """Map one series resolution into audit findings (Tier-1 checks)."""
    series_id = resolved["series_id"]
    findings: list[AuditFinding] = []
    bind_failures = 0
    for issue in resolved["issues"]:
        severity = issue["level"]
        code = str(issue["code"])
        if code == "bind_resolution_failed":
            bind_failures += 1
        findings.append(
            AuditFinding(
                severity=severity,
                code=code,
                series_id=series_id,
                direction=direction,
                message=str(issue["message"]),
                address=issue.get("address"),
            )
        )

    leaf_count = len(resolved["leaves"])
    if series is not None and _is_public_api_series(series, direction) and leaf_count == 0:
        findings.append(
            AuditFinding(
                severity="warning",
                code="empty_public_series",
                series_id=series_id,
                direction=direction,
                message=(
                    f"No resolved {direction} cells for public series "
                    "(codegen skips emission when resolution is otherwise ok)"
                ),
            )
        )

    if leaf_count > 0 and bind_failures > 0:
        findings.append(
            AuditFinding(
                severity="error",
                code="partial_bind_failure",
                series_id=series_id,
                direction=direction,
                message=(
                    f"Resolved {leaf_count} leaf(ves) but {bind_failures} "
                    "bind_resolution_failed issue(s); codegen requires ok=True"
                ),
            )
        )

    if not resolved["ok"] and not any(finding.severity == "error" for finding in findings):
        findings.append(
            AuditFinding(
                severity="error",
                code="resolution_not_ok",
                series_id=series_id,
                direction=direction,
                message="Series resolution reported ok=False without error issues",
            )
        )
    return findings


def _claim_formula_cells(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook_path: Path | str,
) -> dict[str, list[tuple[str, BindingDirection]]]:
    owners: dict[str, list[tuple[str, BindingDirection]]] = {}
    for series in bindings.get("series", []):
        if not isinstance(series, dict):
            continue
        direction = binding_direction(series)
        if direction not in _FORMULA_DIRECTIONS:
            continue
        series_id = str(series.get("id", ""))
        if not series_id:
            continue
        addresses = _expand_series_addresses_for_graph(graph, series, workbook_path=workbook_path)
        if not addresses:
            continue
        typed_direction: BindingDirection = direction
        for address in addresses:
            node = graph.get_node(address)
            if node is None or not node.has_formula:
                continue
            owners.setdefault(address, []).append((series_id, typed_direction))
    return owners


def _duplicate_formula_findings(
    owners: dict[str, list[tuple[str, BindingDirection]]],
) -> list[AuditFinding]:
    findings: list[AuditFinding] = []
    for address, claimed in sorted(owners.items()):
        unique_ids = tuple(dict.fromkeys(series_id for series_id, _direction in claimed))
        if len(unique_ids) < 2:
            continue
        directions = {direction for _series_id, direction in claimed}
        if directions == {"internal"}:
            code = "duplicate_internal_cell_binding"
            finding_direction: BindingDirection = "internal"
        else:
            code = "duplicate_formula_cell_binding"
            finding_direction = claimed[0][1]
        findings.append(
            AuditFinding(
                severity="error",
                code=code,
                series_id=unique_ids[0],
                direction=finding_direction,
                message=(
                    "Formula cell is bound by multiple "
                    f"{'internal ' if code == 'duplicate_internal_cell_binding' else ''}"
                    f"series: {list(unique_ids)}"
                ),
                address=address,
            )
        )
    return findings


def find_duplicate_internal_formula_cell_bindings(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook_path: Path | str,
) -> list[AuditFinding]:
    """Flag formula cells claimed by more than one internal series binding."""
    owners = _claim_formula_cells(graph, bindings, workbook_path=workbook_path)
    internal_only = {
        address: claimed
        for address, claimed in owners.items()
        if {direction for _series_id, direction in claimed} == {"internal"}
    }
    return _duplicate_formula_findings(internal_only)


def find_duplicate_formula_cell_bindings(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook_path: Path | str,
) -> list[AuditFinding]:
    """Flag formula cells claimed by more than one output or internal series.

    Unique-address ownership forbids an internal matrix from covering cells
    already owned by a public output series (and the reverse). Two internal
    series on one formula cell still use `duplicate_internal_cell_binding`.
    """
    owners = _claim_formula_cells(graph, bindings, workbook_path=workbook_path)
    return _duplicate_formula_findings(owners)


def audit_binding_resolutions(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    directions: Sequence[BindingDirection] = DIRECTIONS,
) -> BindingResolutionAuditReport:
    """Resolve bindings for each direction and return a Tier-1 audit report."""
    series_index = _series_by_id(bindings)
    findings: list[AuditFinding] = []
    seen_sparse: set[str] = set()
    workbook_path = Path(workbook)
    loaded_workbook = fastpyxl.load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        for direction in directions:
            report = resolve_series_bindings(
                graph,
                bindings,
                workbook=workbook_path,
                direction=direction,
            )
            for resolved in report["series"]:
                series = series_index.get(resolved["series_id"])
                findings.extend(
                    findings_from_resolution(
                        resolved,
                        direction=direction,
                        series=series,
                    )
                )
                if series is None:
                    continue
                series_id = resolved["series_id"]
                if series_id in seen_sparse:
                    continue
                if not unfilled_label_binds(series):
                    seen_sparse.add(series_id)
                    continue
                seen_sparse.add(series_id)
                findings.extend(
                    find_sparse_label_bind_issues(
                        graph,
                        series,
                        workbook_path=workbook_path,
                        direction=direction,
                        workbook=loaded_workbook,
                    )
                )
    finally:
        loaded_workbook.close()

    if any(direction in _FORMULA_DIRECTIONS for direction in directions):
        findings.extend(
            find_duplicate_formula_cell_bindings(
                graph,
                bindings,
                workbook_path=workbook_path,
            )
        )

    return BindingResolutionAuditReport(findings=tuple(findings))


def format_audit_findings(findings: Iterable[AuditFinding]) -> list[str]:
    """Render findings as stable, human-readable worklist lines."""
    lines: list[str] = []
    ordered = sorted(
        findings,
        key=lambda finding: (
            0 if finding.severity == "error" else 1,
            finding.direction,
            finding.series_id,
            finding.code,
            finding.message,
        ),
    )
    for finding in ordered:
        label = finding.severity.upper()
        address_suffix = f" @ {finding.address}" if finding.address else ""
        lines.append(
            f"{label}  [{finding.direction}]  {finding.series_id}  {finding.code}{address_suffix}"
        )
        lines.append(f"  {finding.message}")
    return lines
