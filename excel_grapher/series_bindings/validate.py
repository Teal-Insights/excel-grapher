from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import has_input_direction, input_mode
from excel_grapher.series_bindings.ranges import expand_data_range_for_graph
from excel_grapher.series_bindings.types import (
    ValidationIssue,
    ValidationReport,
    WorkbookSeriesBindings,
    make_issue,
)
from excel_grapher.series_bindings.versions import IMPLEMENTED_BIND_KINDS, IMPLEMENTED_LAYOUTS

_KNOWN_BIND_KINDS = IMPLEMENTED_BIND_KINDS


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


def _is_formula_graph_node(graph: DependencyGraph, address: str) -> bool:
    node = graph.get_node(address) if address in graph else None
    return bool(node is not None and node.formula is not None)


def _validate_input_binding_overlap(
    graph: DependencyGraph,
    series: dict[str, Any],
    addresses: list[str],
) -> list[ValidationIssue]:
    """Validate leaf vs override semantics for input binding data ranges."""
    issues: list[ValidationIssue] = []
    if not has_input_direction(series):
        return issues

    series_id = str(series.get("id", ""))
    mode = input_mode(series)
    graph_addresses = [address for address in addresses if address in graph]
    non_leaf_graph_addresses = [
        address for address in graph_addresses if not _is_graph_leaf(graph, address)
    ]
    formula_graph_addresses = [
        address for address in graph_addresses if _is_formula_graph_node(graph, address)
    ]

    if mode == "leaf" and non_leaf_graph_addresses:
        issues.append(
            _issue(
                "error",
                "non_leaf_input_overlap",
                "data_range includes non-leaf graph cells; declare input.mode: override "
                "for user-editable formula cells",
                series_id=series_id,
            )
        )
    elif mode == "override" and not formula_graph_addresses:
        issues.append(
            _issue(
                "error",
                "no_formula_override_targets",
                "input.mode override requires at least one formula cell in data_range",
                series_id=series_id,
            )
        )
    return issues


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
    return _validate_bind_geometry(bind, series_id=series_id, context=context)


def _validate_bind_geometry(
    bind: dict[str, Any], *, series_id: str, context: str
) -> list[ValidationIssue]:
    """Statically check skip/include specs and value_map axis and overlap rules."""
    from excel_grapher.series_bindings.geometry import (
        expand_column_specs,
        expand_row_specs,
        parse_value_map,
    )

    issues: list[ValidationIssue] = []
    kind = bind.get("kind")

    if bind.get("skip") is not None and bind.get("include") is not None:
        issues.append(
            _issue(
                "error",
                "invalid_bind_geometry",
                f"{context}: skip and include are mutually exclusive",
                series_id=series_id,
            )
        )

    expand = expand_row_specs if kind == "row_label" else expand_column_specs
    if kind in {"row_label", "column_header"}:
        for field in ("skip", "include"):
            specs = bind.get(field)
            if specs is None:
                continue
            try:
                expand(specs)
            except ValueError as exc:
                issues.append(
                    _issue(
                        "error",
                        "invalid_bind_geometry",
                        f"{context}: {exc}",
                        series_id=series_id,
                    )
                )

    if kind == "value_map":
        values = bind.get("values")
        key_types = {type(key) for key in values} if isinstance(values, dict) else set()
        if len(key_types) > 1:
            issues.append(
                _issue(
                    "error",
                    "invalid_bind_geometry",
                    f"{context}: value_map keys must share one scalar type",
                    series_id=series_id,
                )
            )
        try:
            parse_value_map(values or {})
        except (ValueError, TypeError) as exc:
            issues.append(
                _issue(
                    "error",
                    "invalid_bind_geometry",
                    f"{context}: {exc}",
                    series_id=series_id,
                )
            )

    return issues


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

    exclude_rows = series.get("exclude_rows")
    if exclude_rows is not None:
        from excel_grapher.series_bindings.geometry import expand_row_specs

        try:
            expand_row_specs(exclude_rows)
        except ValueError as exc:
            issues.append(
                _issue(
                    "error",
                    "invalid_bind_geometry",
                    f"exclude_rows: {exc}",
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


def _cell_scoped_dimension_count(series: dict[str, Any]) -> int:
    structure = series.get("structure") or {}
    dimensions = structure.get("dimensions") or []
    return sum(1 for dim in dimensions if isinstance(dim, dict) and dim.get("scope") == "cell")


def _validate_layout_intent(series: dict[str, Any]) -> list[ValidationIssue]:
    """Validate layout-specific structural intent when ``layout`` is declared."""
    issues: list[ValidationIssue] = []
    series_id = str(series.get("id", ""))
    layout = series.get("layout")
    if not isinstance(layout, str):
        return issues

    structure = series.get("structure") or {}
    dimensions = structure.get("dimensions") or []
    dimension_count = len([dim for dim in dimensions if isinstance(dim, dict)])
    cell_scoped_count = _cell_scoped_dimension_count(series)

    if layout == "matrix":
        if dimension_count < 2:
            issues.append(
                _issue(
                    "error",
                    "layout_constraint_violation",
                    "layout 'matrix' requires at least two structure.dimensions entries",
                    series_id=series_id,
                )
            )
        if cell_scoped_count < 1:
            issues.append(
                _issue(
                    "error",
                    "layout_constraint_violation",
                    "layout 'matrix' requires at least one cell-scoped dimension",
                    series_id=series_id,
                )
            )
    elif layout == "series" and cell_scoped_count < 1:
        issues.append(
            _issue(
                "error",
                "layout_constraint_violation",
                "layout 'series' requires at least one cell-scoped dimension",
                series_id=series_id,
            )
        )

    return issues


def _validate_implementation_support(series: dict[str, Any]) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    series_id = str(series.get("id", ""))
    layout = series.get("layout")
    if isinstance(layout, str) and layout not in IMPLEMENTED_LAYOUTS:
        issues.append(
            _issue(
                "error",
                "unknown_layout",
                f"unknown layout {layout!r}",
                series_id=series_id,
            )
        )
    return issues


def _concept_dtype_map(bindings: WorkbookSeriesBindings) -> dict[str, str]:
    scheme = bindings.get("concept_scheme") or {}
    concepts = scheme.get("concepts") or []
    result: dict[str, str] = {}
    for concept in concepts:
        if isinstance(concept, dict) and concept.get("id") and concept.get("dtype") is not None:
            result[str(concept["id"])] = str(concept["dtype"])
    return result


def _read_matches_dtype(read: str, dtype: str) -> bool:
    if read == "auto":
        return True
    if read == dtype:
        return True
    return read == "number" and dtype in {"int", "float", "number"}


def _validate_dtype_read_consistency(
    series: dict[str, Any],
    *,
    concept_dtypes: dict[str, str],
) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    series_id = str(series.get("id", ""))
    structure = series.get("structure") or {}
    measure = structure.get("measure")
    if isinstance(measure, dict):
        measure_dtype = measure.get("dtype")
        bind = measure.get("bind")
        if isinstance(bind, dict) and measure_dtype is not None:
            read = str(bind.get("read", "auto"))
            dtype = str(measure_dtype)
            if not _read_matches_dtype(read, dtype):
                issues.append(
                    _issue(
                        "warning",
                        "dtype_read_mismatch",
                        f"measure dtype {dtype!r} does not match bind read {read!r}",
                        series_id=series_id,
                    )
                )

    for index, dim in enumerate(structure.get("dimensions") or []):
        if not isinstance(dim, dict):
            continue
        concept = str(dim.get("concept", f"dimensions[{index}]"))
        dtype = concept_dtypes.get(concept)
        bind = dim.get("bind")
        if not isinstance(bind, dict) or dtype is None:
            continue
        read = str(bind.get("read", "auto"))
        if not _read_matches_dtype(read, dtype):
            issues.append(
                _issue(
                    "warning",
                    "dtype_read_mismatch",
                    f"dimension {concept!r} dtype {dtype!r} does not match bind read {read!r}",
                    series_id=series_id,
                )
            )

    for index, attr in enumerate(structure.get("attributes") or []):
        if not isinstance(attr, dict):
            continue
        concept = str(attr.get("concept", f"attributes[{index}]"))
        dtype = concept_dtypes.get(concept)
        bind = attr.get("bind")
        if not isinstance(bind, dict) or dtype is None:
            continue
        read = str(bind.get("read", "auto"))
        if not _read_matches_dtype(read, dtype):
            issues.append(
                _issue(
                    "warning",
                    "dtype_read_mismatch",
                    f"attribute {concept!r} dtype {dtype!r} does not match bind read {read!r}",
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
    concept_dtypes = _concept_dtype_map(bindings)

    for series in bindings.get("series", []):
        if not isinstance(series, dict):
            continue
        series_id = str(series.get("id", ""))
        issues.extend(_validate_series_structure(series))
        issues.extend(_validate_layout_intent(series))
        issues.extend(_validate_implementation_support(series))
        issues.extend(_validate_dtype_read_consistency(series, concept_dtypes=concept_dtypes))

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

        issues.extend(_validate_input_binding_overlap(graph, series, addresses))

        graph_leaf_addresses = [
            address
            for address in addresses
            if not intersect_leaves or _is_graph_leaf(graph, address)
        ]
        if input_mode(series) == "override":
            graph_leaf_addresses = [address for address in addresses if address in graph]

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

                resolved = resolve_series_binding(graph, workbook, series, direction="input")
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
