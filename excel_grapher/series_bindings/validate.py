from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.issues import make_issue
from excel_grapher.series_bindings.ranges import expand_data_range_for_graph
from excel_grapher.series_bindings.types import (
    ValidationIssue,
    ValidationReport,
    WorkbookSeriesBindings,
)
from excel_grapher.series_bindings.versions import (
    IMPLEMENTED_BIND_KINDS,
    IMPLEMENTED_LAYOUTS,
    PLANNED_BIND_KINDS,
    PLANNED_LAYOUTS,
)

_KNOWN_BIND_KINDS = IMPLEMENTED_BIND_KINDS | PLANNED_BIND_KINDS


def _issue(
    level: Literal["error", "warning"],
    code: str,
    message: str,
    *,
    series_id: str | None = None,
    address: str | None = None,
) -> ValidationIssue:
    return make_issue(
        level,
        code,
        message,
        series_id=series_id,
        address=address,
    )


def _series_validation_flags(series: dict[str, Any]) -> tuple[bool, bool]:
    validation = series.get("validation") or {}
    intersect = validation.get("intersect_graph_leaves", True)
    unique_key = validation.get("require_unique_key", True)
    return bool(intersect), bool(unique_key)


def _dimension_concepts(series: dict[str, Any]) -> set[str]:
    structure = series.get("structure") or {}
    dims = structure.get("dimensions") or []
    return {str(d.get("concept")) for d in dims if isinstance(d, dict) and d.get("concept")}


def _is_graph_leaf(graph: DependencyGraph, address: str) -> bool:
    node = graph.get_node(address) if address in graph else None
    return bool(node is not None and node.is_leaf)


def _validate_bind_smoke(bind: Any, *, series_id: str, context: str) -> list[ValidationIssue]:
    if not isinstance(bind, dict):
        return [
            _issue(
                "error",
                "invalid_bind",
                f"{context}: bind must be a mapping",
                series_id=series_id,
            )
        ]
    kind = bind.get("kind")
    if kind not in _KNOWN_BIND_KINDS:
        return [
            _issue(
                "error",
                "unknown_bind_kind",
                f"{context}: unknown bind kind {kind!r}",
                series_id=series_id,
            )
        ]
    return []


def _validate_series_structure(series: dict[str, Any]) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    series_id = str(series.get("id", ""))
    sheet = series.get("sheet")
    data_range = series.get("data_range")

    if isinstance(data_range, str) and isinstance(sheet, str) and "!" in str(data_range):
        from excel_grapher.core.address_keys import parse_address
        from excel_grapher.grapher.target_expansion import split_range_target_on_colon

        split = split_range_target_on_colon(data_range)
        start = split[0] if split is not None else data_range
        range_sheet, _ = parse_address(start)
        if range_sheet != sheet:
            issues.append(
                _issue(
                    "error",
                    "sheet_mismatch",
                    f"series sheet {sheet!r} does not match data_range sheet {range_sheet!r}",
                    series_id=series_id,
                )
            )

    structure = series.get("structure")
    if not isinstance(structure, dict):
        return issues

    measure = structure.get("measure")
    if isinstance(measure, dict):
        issues.extend(
            _validate_bind_smoke(measure.get("bind"), series_id=series_id, context="measure")
        )

    for index, dim in enumerate(structure.get("dimensions") or []):
        if not isinstance(dim, dict):
            continue
        bind = dim.get("bind")
        concept = dim.get("concept", f"dimensions[{index}]")
        issues.extend(
            _validate_bind_smoke(
                bind,
                series_id=series_id,
                context=f"dimension {concept!r}",
            )
        )

    for index, attr in enumerate(structure.get("attributes") or []):
        if not isinstance(attr, dict):
            continue
        if "bind" in attr:
            concept = attr.get("concept", f"attributes[{index}]")
            issues.extend(
                _validate_bind_smoke(
                    attr.get("bind"),
                    series_id=series_id,
                    context=f"attribute {concept!r}",
                )
            )

    key = series.get("key")
    if isinstance(key, list):
        dim_concepts = _dimension_concepts(series)
        for concept in key:
            if concept not in dim_concepts:
                issues.append(
                    _issue(
                        "error",
                        "key_not_in_dimensions",
                        f"key concept {concept!r} is not declared in structure.dimensions",
                        series_id=series_id,
                    )
                )

    return issues


def _validate_implementation_support(series: dict[str, Any]) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    series_id = str(series.get("id", ""))
    layout = series.get("layout")
    if isinstance(layout, str) and layout in PLANNED_LAYOUTS:
        issues.append(
            _issue(
                "warning",
                "layout_not_implemented",
                f"layout {layout!r} is defined in schema 1.1.0 but not yet supported by resolve/codegen",
                series_id=series_id,
            )
        )
    structure = series.get("structure") or {}
    for dim in structure.get("dimensions") or []:
        if not isinstance(dim, dict):
            continue
        bind = dim.get("bind")
        if isinstance(bind, dict):
            kind = bind.get("kind")
            if kind in PLANNED_BIND_KINDS:
                issues.append(
                    _issue(
                        "warning",
                        "bind_not_implemented",
                        f"bind kind {kind!r} is not yet supported by resolve/codegen",
                        series_id=series_id,
                    )
                )
    for attr in structure.get("attributes") or []:
        if isinstance(attr, dict) and isinstance(attr.get("bind"), dict):
            kind = attr["bind"].get("kind")
            if kind in PLANNED_BIND_KINDS:
                issues.append(
                    _issue(
                        "warning",
                        "bind_not_implemented",
                        f"bind kind {kind!r} is not yet supported by resolve/codegen",
                        series_id=series_id,
                    )
                )
    measure = structure.get("measure")
    if isinstance(measure, dict) and isinstance(measure.get("bind"), dict):
        kind = measure["bind"].get("kind")
        if kind in PLANNED_BIND_KINDS:
            issues.append(
                _issue(
                    "warning",
                    "bind_not_implemented",
                    f"bind kind {kind!r} is not yet supported by resolve/codegen",
                    series_id=series_id,
                )
            )
    if (
        isinstance(layout, str)
        and layout not in IMPLEMENTED_LAYOUTS
        and layout not in PLANNED_LAYOUTS
    ):
        issues.append(
            _issue(
                "error",
                "unknown_layout",
                f"unknown layout {layout!r}",
                series_id=series_id,
            )
        )
    return issues


def validate_series_bindings(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str | None = None,
) -> ValidationReport:
    """Validate binding manifests against an extracted dependency graph."""
    issues: list[ValidationIssue] = []

    for series in bindings.get("series", []):
        if not isinstance(series, dict):
            continue
        series_id = str(series.get("id", ""))
        issues.extend(_validate_series_structure(series))
        issues.extend(_validate_implementation_support(series))

        data_range = series.get("data_range")
        if not isinstance(data_range, str):
            continue

        intersect_leaves, require_unique_key = _series_validation_flags(series)
        try:
            addresses = expand_data_range_for_graph(
                graph,
                data_range,
                workbook=workbook,
            )
        except (ValueError, TypeError) as exc:
            issues.append(
                _issue(
                    "error",
                    "invalid_data_range",
                    str(exc),
                    series_id=series_id,
                )
            )
            continue

        if not addresses:
            issues.append(
                _issue(
                    "error",
                    "empty_data_range",
                    "data_range expands to zero cells",
                    series_id=series_id,
                )
            )
            continue

        graph_leaf_addresses = [
            address
            for address in addresses
            if not intersect_leaves or _is_graph_leaf(graph, address)
        ]

        if require_unique_key:
            cell_scoped_keys = [
                str(c)
                for c in (series.get("key") or [])
                if c in _dimension_concepts(series)
                and any(
                    isinstance(d, dict) and d.get("concept") == c and d.get("scope") == "cell"
                    for d in (series.get("structure") or {}).get("dimensions") or []
                )
            ]
            if not graph_leaf_addresses:
                continue
            if cell_scoped_keys and workbook is None:
                issues.append(
                    _issue(
                        "warning",
                        "unique_key_deferred",
                        "require_unique_key is set but coordinate resolution is not implemented yet "
                        f"(cell-scoped key dimensions: {cell_scoped_keys})",
                        series_id=series_id,
                    )
                )
            elif workbook is not None:
                from excel_grapher.series_bindings.resolve import resolve_series_binding

                resolved = resolve_series_binding(graph, workbook, series)
                issues.extend(resolved["issues"])
                if resolved["requires_address"]:
                    issues.append(
                        _issue(
                            "warning",
                            "requires_address",
                            "Duplicate or ambiguous record keys require address disambiguation",
                            series_id=series_id,
                        )
                    )

    ok = not any(i["level"] == "error" for i in issues)
    return {"ok": ok, "issues": issues}
