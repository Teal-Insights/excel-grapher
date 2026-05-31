"""Resolve series binding cells to coordinate maps for setter and compute codegen."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Literal

import fastpyxl
import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_key, parse_address, quote_sheet_if_needed
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.issues import make_issue
from excel_grapher.series_bindings.normalize import has_input_direction, has_output_direction
from excel_grapher.series_bindings.ranges import expand_data_range_for_graph
from excel_grapher.series_bindings.types import (
    LeafResolution,
    ResolutionIssue,
    ResolutionReport,
    Scalar,
    SeriesResolution,
    WorkbookSeriesBindings,
)

BindingDirection = Literal["input", "output"]

_TRAILING_UNIT_RE = re.compile(r"\s*\([^)]*\)\s*$")


class _WorkbookValues:
    """Lazy cached value reader for bind cells outside the dependency graph."""

    def __init__(self, path: Path | str) -> None:
        self._path = Path(path)
        self._workbook_cache: fastpyxl.Workbook | None = None

    def _workbook(self) -> fastpyxl.Workbook:
        if self._workbook_cache is not None:
            return self._workbook_cache
        keep_vba = self._path.suffix.lower() == ".xlsm"
        self._workbook_cache = fastpyxl.load_workbook(
            self._path,
            data_only=True,
            read_only=True,
            keep_vba=keep_vba,
        )
        return self._workbook_cache

    def read(self, address: str) -> Any:
        sheet, coord = parse_address(address)
        wb = self._workbook()
        if sheet not in wb.sheetnames:
            raise KeyError(f"Sheet {sheet!r} not found in workbook")
        return wb[sheet][coord].value


def _read_cell_value(graph: DependencyGraph, reader: _WorkbookValues, address: str) -> Any:
    if address in graph:
        node = graph.get_node(address)
        if node is not None:
            return node.value
    return reader.read(address)


def _normalize_string(value: str, normalize: str) -> str:
    if normalize == "none":
        return value
    if normalize == "strip":
        return value.strip()
    if normalize == "strip_trailing_unit":
        return _TRAILING_UNIT_RE.sub("", value).strip()
    return value


def _coerce_value(raw: Any, read_as: str) -> Scalar:
    if raw is None:
        return None
    if read_as == "auto":
        if isinstance(raw, bool):
            return raw
        if isinstance(raw, int) and not isinstance(raw, bool):
            return raw
        if isinstance(raw, float):
            return raw
        if isinstance(raw, str):
            return raw
        return str(raw)
    if read_as == "string":
        return str(raw)
    if read_as == "int":
        return int(raw)
    if read_as == "float":
        return float(raw)
    if read_as == "number":
        if isinstance(raw, int) and not isinstance(raw, bool):
            return raw
        return float(raw)
    if read_as == "bool":
        if isinstance(raw, bool):
            return raw
        if isinstance(raw, (int, float)):
            return bool(raw)
        text = str(raw).strip().lower()
        if text in {"true", "1", "yes"}:
            return True
        if text in {"false", "0", "no"}:
            return False
        raise ValueError(f"Cannot coerce {raw!r} to bool")
    raise ValueError(f"Unknown read mode: {read_as!r}")


def _parse_data_cell(address: str) -> tuple[str, str, int]:
    sheet, coord = parse_address(address)
    col, row = fastpyxl.utils.cell.coordinate_from_string(coord)
    return sheet, col, row


def _execute_bind(
    bind: dict[str, Any],
    *,
    graph: DependencyGraph,
    reader: _WorkbookValues,
    data_address: str,
) -> Scalar:
    kind = bind.get("kind")
    read_as = str(bind.get("read", "auto"))
    normalize = str(bind.get("normalize", "strip"))

    if kind == "data_cell":
        raw = _read_cell_value(graph, reader, data_address)
        return _coerce_value(raw, read_as)

    if kind == "constant":
        return bind.get("value")

    if kind == "cell":
        address = str(bind["address"])
        raw = _read_cell_value(graph, reader, address)
        if read_as in {"auto", "string"} and isinstance(raw, str):
            return _normalize_string(raw, normalize)
        return _coerce_value(raw, read_as)

    sheet, col, row = _parse_data_cell(data_address)
    quoted_sheet = quote_sheet_if_needed(sheet)

    if kind == "column_header":
        header_row = int(bind["header_row"])
        header_address = format_key(quoted_sheet, f"{col}{header_row}")
        raw = _read_cell_value(graph, reader, header_address)
        if read_as in {"auto", "string"} and isinstance(raw, str):
            return _normalize_string(raw, normalize)
        return _coerce_value(raw, read_as)

    if kind == "row_hierarchy":
        raise ValueError(
            "bind kind 'row_hierarchy' is defined in schema 1.1.0 but not yet implemented"
        )

    if kind == "row_label":
        label_column = str(bind["label_column"])
        label_address = format_key(quoted_sheet, f"{label_column}{row}")
        raw = _read_cell_value(graph, reader, label_address)
        if read_as in {"auto", "string"} and isinstance(raw, str):
            return _normalize_string(raw, normalize)
        return _coerce_value(raw, read_as)

    raise ValueError(f"Unknown bind kind: {kind!r}")


def _include_in_record(field: dict[str, Any], default: bool) -> bool:
    if "include_in_record" in field:
        return bool(field["include_in_record"])
    return default


def _build_input_record(
    *,
    coordinates: dict[str, Scalar],
    series: dict[str, Any],
    measure_concept: str,
) -> dict[str, Scalar]:
    record: dict[str, Scalar] = {}
    key_concepts = [str(c) for c in (series.get("key") or [])]

    for concept in key_concepts:
        if concept in coordinates:
            record[concept] = coordinates[concept]

    obs_value = coordinates.get(measure_concept)
    if obs_value is not None or measure_concept in coordinates:
        record[measure_concept] = obs_value

    for concept, value in (series.get("series_context") or {}).items():
        record[str(concept)] = value

    structure = series.get("structure") or {}
    for dim in structure.get("dimensions") or []:
        if not isinstance(dim, dict):
            continue
        concept = str(dim.get("concept", ""))
        if _include_in_record(dim, default=True) and concept in coordinates:
            record[concept] = coordinates[concept]

    for attr in structure.get("attributes") or []:
        if not isinstance(attr, dict):
            continue
        concept = str(attr.get("concept", ""))
        if not _include_in_record(attr, default=False):
            continue
        if "value" in attr:
            record[concept] = attr["value"]
        elif concept in coordinates:
            record[concept] = coordinates[concept]

    return record


def _build_output_record(
    *,
    coordinates: dict[str, Scalar],
    series: dict[str, Any],
    measure_concept: str,
) -> dict[str, Scalar]:
    """Build a record with all declared dimensions and attributes (OBS_VALUE filled at runtime)."""
    record: dict[str, Scalar] = {}

    for concept, value in (series.get("series_context") or {}).items():
        record[str(concept)] = value

    structure = series.get("structure") or {}
    for dim in structure.get("dimensions") or []:
        if not isinstance(dim, dict):
            continue
        concept = str(dim.get("concept", ""))
        if concept in coordinates:
            record[concept] = coordinates[concept]

    for attr in structure.get("attributes") or []:
        if not isinstance(attr, dict):
            continue
        concept = str(attr.get("concept", ""))
        if "value" in attr:
            record[concept] = attr["value"]
        elif concept in coordinates:
            record[concept] = coordinates[concept]

    if measure_concept not in record:
        record[measure_concept] = coordinates.get(measure_concept)

    return record


def _extract_key(coordinates: dict[str, Scalar], key_concepts: list[str]) -> dict[str, Scalar]:
    return {concept: coordinates[concept] for concept in key_concepts if concept in coordinates}


def _is_graph_leaf(graph: DependencyGraph, address: str) -> bool:
    node = graph.get_node(address) if address in graph else None
    return bool(node is not None and node.is_leaf)


def _is_graph_node(graph: DependencyGraph, address: str) -> bool:
    return address in graph


def _select_addresses(
    graph: DependencyGraph,
    expanded_addresses: list[str],
    *,
    direction: BindingDirection,
    validation: dict[str, Any],
    series_id: str,
) -> tuple[list[str], list[ResolutionIssue]]:
    issues: list[ResolutionIssue] = []
    if direction == "input":
        intersect = bool(validation.get("intersect_graph_leaves", True))
        if not intersect:
            return expanded_addresses, issues
        selected = [address for address in expanded_addresses if _is_graph_leaf(graph, address)]
    else:
        intersect = bool(validation.get("intersect_graph_nodes", True))
        if not intersect:
            return expanded_addresses, issues
        selected = [address for address in expanded_addresses if _is_graph_node(graph, address)]

    skipped = len(expanded_addresses) - len(selected)
    if skipped > 0 and bool(validation.get("warn_on_partial_overlap", True)):
        issues.append(
            make_issue(
                "warning",
                "partial_graph_overlap",
                f"Skipped {skipped} cell(s) in data_range not present in graph for {direction} binding",
                series_id=series_id,
            )
        )
    return selected, issues


def resolve_series_binding(
    graph: DependencyGraph,
    workbook: Path | str,
    series: dict[str, Any],
    *,
    direction: BindingDirection = "input",
) -> SeriesResolution:
    """Resolve each participating cell in a series binding to coordinates and record fields."""
    series_id = str(series.get("id", ""))
    issues: list[ResolutionIssue] = []
    leaves: list[LeafResolution] = []
    requires_address = False

    data_range = series.get("data_range")
    if not isinstance(data_range, str):
        return {
            "series_id": series_id,
            "ok": False,
            "requires_address": True,
            "leaves": [],
            "issues": [
                make_issue(
                    "error", "missing_data_range", "series has no data_range", series_id=series_id
                )
            ],
        }

    try:
        expanded_addresses = expand_data_range_for_graph(graph, data_range, workbook=workbook)
    except (ValueError, TypeError) as exc:
        return {
            "series_id": series_id,
            "ok": False,
            "requires_address": True,
            "leaves": [],
            "issues": [make_issue("error", "invalid_data_range", str(exc), series_id=series_id)],
        }

    validation = series.get("validation") or {}
    addresses, overlap_issues = _select_addresses(
        graph,
        expanded_addresses,
        direction=direction,
        validation=validation,
        series_id=series_id,
    )
    issues.extend(overlap_issues)
    if not addresses:
        issues.append(
            make_issue(
                "warning",
                "no_resolved_cells",
                f"No resolved {direction} cells in data_range after graph intersection",
                series_id=series_id,
            )
        )

    structure = series.get("structure") or {}
    measure = structure.get("measure") or {}
    measure_concept = str(measure.get("concept", "OBS_VALUE"))
    measure_bind = measure.get("bind") or {"kind": "data_cell"}
    key_concepts = [str(c) for c in (series.get("key") or [])]
    require_unique_key = bool(validation.get("require_unique_key", True))

    reader = _WorkbookValues(workbook)
    series_coordinates: dict[str, Scalar] = {}
    seen_keys: dict[tuple[tuple[str, Scalar], ...], str] = {}

    build_record = _build_input_record if direction == "input" else _build_output_record

    for address in addresses:
        coordinates: dict[str, Scalar] = {}
        try:
            if direction == "input":
                coordinates[measure_concept] = _execute_bind(
                    measure_bind if isinstance(measure_bind, dict) else {"kind": "data_cell"},
                    graph=graph,
                    reader=reader,
                    data_address=address,
                )

            for dim in structure.get("dimensions") or []:
                if not isinstance(dim, dict):
                    continue
                concept = str(dim.get("concept", ""))
                bind = dim.get("bind")
                if not isinstance(bind, dict):
                    continue
                scope = dim.get("scope")
                if scope == "series" and concept in series_coordinates:
                    coordinates[concept] = series_coordinates[concept]
                    continue
                value = _execute_bind(
                    bind,
                    graph=graph,
                    reader=reader,
                    data_address=address,
                )
                coordinates[concept] = value
                if scope == "series":
                    series_coordinates[concept] = value

            for attr in structure.get("attributes") or []:
                if not isinstance(attr, dict):
                    continue
                concept = str(attr.get("concept", ""))
                if "value" in attr:
                    coordinates[concept] = attr["value"]
                elif "bind" in attr and isinstance(attr["bind"], dict):
                    coordinates[concept] = _execute_bind(
                        attr["bind"],
                        graph=graph,
                        reader=reader,
                        data_address=address,
                    )
        except (KeyError, ValueError, TypeError) as exc:
            issues.append(
                make_issue(
                    "error",
                    "bind_resolution_failed",
                    str(exc),
                    series_id=series_id,
                    address=address,
                )
            )
            continue

        key = _extract_key(coordinates, key_concepts)
        record = build_record(
            coordinates=coordinates,
            series=series,
            measure_concept=measure_concept,
        )
        leaves.append(
            {
                "address": address,
                "coordinates": coordinates,
                "key": key,
                "record": record,
            }
        )

        if require_unique_key and key_concepts:
            key_tuple = tuple(sorted(key.items()))
            if key_tuple in seen_keys:
                requires_address = True
                issues.append(
                    make_issue(
                        "error",
                        "duplicate_key",
                        f"Duplicate key {dict(key_tuple)!r} at {address} "
                        f"(first seen at {seen_keys[key_tuple]})",
                        series_id=series_id,
                        address=address,
                    )
                )
            else:
                seen_keys[key_tuple] = address

    layout = series.get("layout")
    if layout == "scalar" and not key_concepts and len(leaves) > 1:
        issues.append(
            make_issue(
                "error",
                "keyless_scalar_ambiguous",
                "Keyless scalar binding must resolve to exactly one leaf; "
                f"got {len(leaves)} leaves in data_range",
                series_id=series_id,
            )
        )

    ok = not any(i["level"] == "error" for i in issues)
    return {
        "series_id": series_id,
        "ok": ok,
        "requires_address": requires_address,
        "leaves": leaves,
        "issues": issues,
    }


def _series_supports_direction(series: dict[str, Any], direction: BindingDirection) -> bool:
    if direction == "input":
        return has_input_direction(series)
    return has_output_direction(series)


def resolve_series_bindings(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str | None = None,
    direction: BindingDirection = "input",
) -> ResolutionReport:
    """Resolve all series in a binding manifest for the given direction."""
    if workbook is None:
        return {
            "ok": False,
            "series": [],
            "issues": [
                make_issue(
                    "error",
                    "missing_workbook",
                    "workbook path is required for resolution",
                )
            ],
        }

    series_results: list[SeriesResolution] = []
    all_issues: list[ResolutionIssue] = []
    for series in bindings.get("series", []):
        if not isinstance(series, dict):
            continue
        if not _series_supports_direction(series, direction):
            continue
        result = resolve_series_binding(graph, workbook, series, direction=direction)
        series_results.append(result)
        all_issues.extend(result["issues"])

    ok = all(r["ok"] for r in series_results) and not any(i["level"] == "error" for i in all_issues)
    return {"ok": ok, "series": series_results, "issues": all_issues}
