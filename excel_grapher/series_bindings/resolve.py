"""Resolve series binding cells to coordinate maps for setter and compute codegen."""

from __future__ import annotations

import re
import warnings
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from typing import Any, Literal

import fastpyxl
import fastpyxl.utils.cell
from fastpyxl.utils import column_index_from_string, get_column_letter

from excel_grapher.core.address_keys import format_key, parse_address
from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.coerce import coerce_constant, coerce_scalar
from excel_grapher.series_bindings.geometry import (
    expand_column_specs,
    expand_row_specs,
    parse_value_map,
)
from excel_grapher.series_bindings.graph_predicates import (
    is_graph_formula_node,
    is_graph_leaf,
    is_graph_node,
)
from excel_grapher.series_bindings.normalize import (
    InputMode,
    component_for_field,
    effective_dimension_id,
    effective_validation,
    has_constant_direction,
    has_input_direction,
    has_internal_direction,
    has_output_direction,
    input_mode,
)
from excel_grapher.series_bindings.ranges import (
    apply_series_excludes,
    expand_series_data_ranges_for_graph,
    series_data_ranges,
)
from excel_grapher.series_bindings.types import (
    LeafResolution,
    ResolutionIssue,
    ResolutionReport,
    Scalar,
    SeriesResolution,
    WorkbookSeriesBindings,
    make_issue,
)

BindingDirection = Literal["input", "output", "internal", "constant"]

_TRAILING_UNIT_RE = re.compile(r"\s*\([^)]*\)\s*$")


class UnknownBindKindError(ValueError):
    """Bind mapping used a `kind` that `_execute_bind` does not implement.

    Attributes:
        kind: The unimplemented bind `kind` value.
        series_id: Binding series id when raised from `resolve_key_domain`.
        address: Data cell that triggered the bind.
        field_name: Key field whose bind used `kind`.
    """

    def __init__(
        self,
        kind: object,
        *,
        series_id: str = "",
        address: str = "",
        field_name: str = "",
    ) -> None:
        self.kind = kind
        self.series_id = series_id
        self.address = address
        self.field_name = field_name
        message = f"Unknown bind kind: {kind!r}"
        if series_id and address and field_name:
            message = (
                f"series {series_id!r} cell {address}: key field "
                f"{field_name!r} bind failed: {message}"
            )
        super().__init__(message)


class PartialKeyDomainError(ValueError):
    """A declared key field resolved for some members and not others.

    Attributes:
        series_id: Binding series id.
        unresolved: Data-cell address to the key fields missing on that cell.
    """

    def __init__(self, series_id: str, unresolved: Mapping[str, Sequence[str]]) -> None:
        self.series_id = series_id
        self.unresolved = {addr: tuple(fields) for addr, fields in unresolved.items()}
        parts = [
            f"{addr} ({', '.join(repr(field) for field in fields)})"
            for addr, fields in self.unresolved.items()
        ]
        super().__init__(
            f"series {series_id!r}: key did not fully resolve for cells {', '.join(parts)}"
        )


class _WorkbookValues:
    """Lazy cached value reader for bind cells outside the dependency graph.

    Opens the workbook in `read_only` mode so unused sheets are never bound.
    Each touched sheet is streamed only through the last requested row; the
    rest of that sheet and every unread sheet stay unparsed.
    """

    def __init__(self, path: Path | str) -> None:
        self._path = Path(path)
        self._workbook_cache: fastpyxl.Workbook | None = None
        self._sheet_values: dict[str, dict[str, Any]] = {}

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

    def prefetch(self, addresses: Iterable[str], *, graph: DependencyGraph | None = None) -> None:
        """Stream each touched sheet only through the last requested row."""
        wanted_by_sheet: dict[str, set[str]] = {}
        for address in addresses:
            if graph is not None and address in graph:
                continue
            sheet, coord = parse_address(address)
            wanted_by_sheet.setdefault(sheet, set()).add(coord)
        for sheet, wanted in wanted_by_sheet.items():
            cached = self._sheet_values.get(sheet)
            if cached is not None and wanted <= cached.keys():
                continue
            values = _stream_sheet_values(self._workbook(), sheet, wanted)
            self._sheet_values.setdefault(sheet, {}).update(values)

    def read(self, address: str) -> Any:
        sheet, coord = parse_address(address)
        cached = self._sheet_values.get(sheet)
        if cached is None or coord not in cached:
            self.prefetch((address,))
            cached = self._sheet_values.get(sheet, {})
        return cached.get(coord)

    def close(self) -> None:
        """Close the cached workbook, if open."""
        if self._workbook_cache is not None:
            self._workbook_cache.close()
            self._workbook_cache = None
        self._sheet_values.clear()

    def __enter__(self) -> _WorkbookValues:
        return self

    def __exit__(self, *args: object) -> None:
        self.close()


def _stream_sheet_values(
    workbook: fastpyxl.Workbook,
    sheet: str,
    wanted: set[str],
) -> dict[str, Any]:
    """Stream one worksheet until every requested coordinate's row is passed."""
    from fastpyxl.worksheet._reader import WorkSheetParser

    if not wanted:
        return {}
    if sheet not in workbook.sheetnames:
        raise KeyError(f"Sheet {sheet!r} not found in workbook")
    worksheet = workbook[sheet]
    get_source = getattr(worksheet, "_get_source", None)
    if get_source is None:
        raise TypeError(f"worksheet {sheet!r} does not support streamed reads")
    max_row = max(fastpyxl.utils.cell.coordinate_from_string(coord)[1] for coord in wanted)
    values: dict[str, Any] = {}
    with get_source() as source:
        parser = WorkSheetParser(
            source,
            getattr(worksheet, "_shared_strings", []),
            data_only=True,
            epoch=workbook.epoch,
            date_formats=workbook._date_formats,
            timedelta_formats=workbook._timedelta_formats,
        )
        for row_idx, cells in parser.parse():
            for row, column, value, _dtype, _style, _cached in cells:
                coord = f"{get_column_letter(column)}{row}"
                if coord in wanted:
                    values[coord] = value
            if row_idx >= max_row:
                break
    for coord in wanted:
        values.setdefault(coord, None)
    return values


def _bind_source_addresses(bind: dict[str, Any], data_address: str) -> list[str]:
    """Return workbook cells `bind` may read for `data_address`."""
    kind = bind.get("kind")
    if kind == "data_cell":
        return [data_address]
    if kind == "cell":
        return [str(bind["address"])]
    if kind in {"constant", "value_map", "sheet_name"}:
        return []
    if kind not in {"column_header", "row_label"}:
        return []
    sheet, col, row = _parse_data_cell(data_address)
    fill = bool(bind.get("fill", False))
    if kind == "column_header":
        header_row = int(bind["header_row"])
        index = column_index_from_string(col)
        sources, is_include = _label_source_indices(bind, axis="columns")
        candidates = range(index, 0, -1) if fill else (index,)
        return [
            format_key(sheet, f"{get_column_letter(candidate)}{header_row}")
            for candidate in candidates
            if _is_label_source(candidate, sources, is_include=is_include)
        ]
    label_column = str(bind["label_column"])
    sources, is_include = _label_source_indices(bind, axis="rows")
    candidates = range(row, 0, -1) if fill else (row,)
    return [
        format_key(sheet, f"{label_column}{candidate}")
        for candidate in candidates
        if _is_label_source(candidate, sources, is_include=is_include)
    ]


def _structure_source_addresses(
    series: dict[str, Any],
    cells: Sequence[str],
    *,
    include_measure: bool = False,
    include_attributes: bool = False,
) -> list[str]:
    """Collect label and header cells for `cells` from `series` structure."""
    structure = series.get("structure") or {}
    addresses: list[str] = []
    if include_measure:
        measure = structure.get("measure") or {}
        measure_bind = measure.get("bind") or {"kind": "data_cell"}
        if isinstance(measure_bind, dict):
            for cell in cells:
                addresses.extend(_bind_source_addresses(measure_bind, cell))
    for dim in structure.get("dimensions") or []:
        if not isinstance(dim, dict):
            continue
        bind = dim.get("bind")
        if not isinstance(bind, dict):
            continue
        for cell in cells:
            addresses.extend(_bind_source_addresses(bind, cell))
    if include_attributes:
        for attr in structure.get("attributes") or []:
            if not isinstance(attr, dict):
                continue
            bind = _attribute_bind(attr)
            if bind is None:
                continue
            for cell in cells:
                addresses.extend(_bind_source_addresses(bind, cell))
    return addresses


def _read_cell_value(graph: DependencyGraph | None, reader: _WorkbookValues, address: str) -> Any:
    if graph is not None and address in graph:
        node = graph.get_node(address)
        if node is not None:
            return node.value
    return reader.read(address)


def _lookup_concept_dtype(
    concept_scheme: dict[str, Any] | None,
    series: dict[str, Any],
    concept_name: str,
) -> str | None:
    measure = (series.get("structure") or {}).get("measure") or {}
    if concept_name == str(measure.get("concept", "OBS_VALUE")):
        measure_dtype = measure.get("dtype")
        if measure_dtype is not None:
            return str(measure_dtype)
    if concept_scheme:
        for concept in concept_scheme.get("concepts") or []:
            if isinstance(concept, dict) and str(concept.get("id")) == concept_name:
                dtype = concept.get("dtype")
                if dtype is not None:
                    return str(dtype)
    return None


def _component_dtype(
    concept_scheme: dict[str, Any] | None,
    series: dict[str, Any],
    component: dict[str, Any],
) -> str | None:
    """Inferred dtype for a dimension or attribute: declared dtype, else concept scheme."""
    dtype = component.get("dtype")
    if dtype is not None:
        return str(dtype)
    return _lookup_concept_dtype(concept_scheme, series, str(component.get("concept", "")))


def _effective_read_as(
    bind: dict[str, Any],
    *,
    inferred_dtype: str | None,
) -> str:
    if "read" in bind:
        return str(bind["read"])
    if inferred_dtype in {"string", "int", "float", "number", "bool", "datetime"}:
        return inferred_dtype
    return "auto"


def _normalize_string(value: str, normalize: str) -> str:
    if normalize == "none":
        return value
    if normalize == "strip":
        return value.strip()
    if normalize == "strip_trailing_unit":
        return _TRAILING_UNIT_RE.sub("", value).strip()
    return value


def _summarize_affected_addresses(addresses: list[str]) -> str:
    if len(addresses) == 1:
        return addresses[0]
    return f"{len(addresses)} cells ({addresses[0]}\u2013{addresses[-1]})"


def _bind_resolution_issues(
    failures: dict[str, list[str]],
    *,
    series_id: str,
) -> list[ResolutionIssue]:
    issues: list[ResolutionIssue] = []
    for message, addresses in failures.items():
        if len(addresses) == 1:
            issues.append(
                make_issue(
                    "error",
                    "bind_resolution_failed",
                    message,
                    series_id=series_id,
                    address=addresses[0],
                )
            )
            continue
        summary = _summarize_affected_addresses(addresses)
        issues.append(
            make_issue(
                "error",
                "bind_resolution_failed",
                f"{message} (affects {summary})",
                series_id=series_id,
                address=None,
            )
        )
    return issues


def _is_blank_label(raw: Any) -> bool:
    return raw is None or (isinstance(raw, str) and not raw.strip())


def _missing_policy(bind: dict[str, Any]) -> str:
    """Return the effective missing-label policy (`missing: null` in YAML is None)."""
    if "missing" not in bind:
        return "error"
    value = bind["missing"]
    return "null" if value is None else str(value)


def _label_source_indices(bind: dict[str, Any], *, axis: str) -> tuple[set[int] | None, bool]:
    """Return (allowed source indices or None for all, whether include was used)."""
    skip = bind.get("skip")
    include = bind.get("include")
    if skip is not None and include is not None:
        raise ValueError("skip and include are mutually exclusive on a bind")
    expand = expand_row_specs if axis == "rows" else expand_column_specs
    if include is not None:
        return expand(include), True
    if skip is not None:
        return expand(skip), False
    return None, False


def _is_label_source(index: int, sources: set[int] | None, *, is_include: bool) -> bool:
    if sources is None:
        return True
    return index in sources if is_include else index not in sources


def _resolve_label(
    bind: dict[str, Any],
    *,
    graph: DependencyGraph | None,
    reader: _WorkbookValues,
    axis: str,
    index: int,
    address_for: Any,
    concept_hint: str,
) -> Any:
    """Read a row_label or column_header source value honoring skip/include/fill.

    Args:
        bind: The bind mapping carrying skip/include/fill/missing fields.
        graph: Dependency graph consulted before the workbook reader.
        reader: Lazy workbook value reader for cells outside the graph.
        axis: `rows` for row_label (walk up) or `columns` for column_header
            (walk left).
        index: 1-based data row (rows axis) or data column (columns axis).
        address_for: Callable mapping a source index to a sheet-qualified
            address.
        concept_hint: Bind description used in missing-label error messages.

    Returns:
        The raw label value, or None when no source label exists and the
        bind's missing policy is `null`.

    Raises:
        ValueError: When no source label exists and the policy is `error`.
    """
    sources, is_include = _label_source_indices(bind, axis=axis)
    fill = bool(bind.get("fill", False))

    def source_value(candidate: int) -> Any:
        if not _is_label_source(candidate, sources, is_include=is_include):
            return None
        raw = _read_cell_value(graph, reader, address_for(candidate))
        return None if _is_blank_label(raw) else raw

    raw = source_value(index)
    if raw is None and fill:
        for candidate in range(index - 1, 0, -1):
            raw = source_value(candidate)
            if raw is not None:
                break

    if raw is None:
        if _missing_policy(bind) == "null":
            return None
        direction = "at or above row" if axis == "rows" else "at or left of column"
        raise ValueError(f"{concept_hint}: no source label {direction} {index}")
    return raw


def _execute_bind(
    bind: dict[str, Any],
    *,
    graph: DependencyGraph | None,
    reader: _WorkbookValues,
    data_address: str,
    inferred_read_as: str | None = None,
) -> Scalar:
    kind = bind.get("kind")
    read_as = _effective_read_as(bind, inferred_dtype=inferred_read_as)
    normalize = str(bind.get("normalize", "strip"))

    if kind == "data_cell":
        raw = _read_cell_value(graph, reader, data_address)
        return coerce_scalar(raw, read_as)

    if kind == "constant":
        return coerce_constant(bind.get("value"), read_as=read_as)

    if kind == "cell":
        address = str(bind["address"])
        raw = _read_cell_value(graph, reader, address)
        if read_as in {"auto", "string"} and isinstance(raw, str):
            return _normalize_string(raw, normalize)
        return coerce_scalar(raw, read_as)

    sheet, col, row = _parse_data_cell(data_address)

    if kind == "column_header":
        header_row = int(bind["header_row"])
        raw = _resolve_label(
            bind,
            graph=graph,
            reader=reader,
            axis="columns",
            index=column_index_from_string(col),
            address_for=lambda c: format_key(sheet, f"{get_column_letter(c)}{header_row}"),
            concept_hint=f"column_header row {header_row}",
        )
        if raw is None:
            return None
        if read_as in {"auto", "string"} and isinstance(raw, str):
            return _normalize_string(raw, normalize)
        return coerce_scalar(raw, read_as)

    if kind == "row_label":
        label_column = str(bind["label_column"])
        raw = _resolve_label(
            bind,
            graph=graph,
            reader=reader,
            axis="rows",
            index=row,
            address_for=lambda r: format_key(sheet, f"{label_column}{r}"),
            concept_hint=f"row_label column {label_column}",
        )
        if raw is None:
            return None
        if read_as in {"auto", "string"} and isinstance(raw, str):
            return _normalize_string(raw, normalize)
        return coerce_scalar(raw, read_as)

    if kind == "value_map":
        axis, mapping = parse_value_map(bind.get("values") or {})
        index = row if axis == "rows" else column_index_from_string(col)
        for value, indices in mapping.items():
            if index in indices:
                return coerce_constant(value, read_as=read_as)
        if _missing_policy(bind) == "null":
            return None
        unit = "row" if axis == "rows" else "column"
        raise ValueError(f"value_map: no value covers data {unit} {index}")

    if kind == "sheet_name":
        values = bind.get("values")
        if isinstance(values, dict) and values:
            if sheet not in values:
                raise ValueError(f"sheet_name: worksheet {sheet!r} is not in values")
            raw = values[sheet]
        else:
            raw = sheet
        if read_as in {"auto", "string"} and isinstance(raw, str):
            return _normalize_string(raw, normalize)
        return coerce_constant(raw, read_as=read_as)

    raise UnknownBindKindError(kind)


def _parse_data_cell(address: str) -> tuple[str, str, int]:
    sheet, coord = parse_address(address)
    col, row = fastpyxl.utils.cell.coordinate_from_string(coord)
    return sheet, col, row


def _include_in_record(field: dict[str, Any], default: bool) -> bool:
    if "include_in_record" in field:
        return bool(field["include_in_record"])
    return default


def _attribute_bind(attr: dict[str, Any]) -> dict[str, Any] | None:
    """Normalize an attribute declaration to a bind mapping for resolution."""
    bind = attr.get("bind")
    if isinstance(bind, dict):
        return bind
    if "value" in attr:
        return {"kind": "constant", "value": attr["value"]}
    return None


def _coerce_series_context(
    series: dict[str, Any],
    *,
    concept_scheme: dict[str, Any] | None,
) -> dict[str, Scalar]:
    """Coerce manifest ``series_context`` values using concept-scheme dtypes."""
    raw = series.get("series_context") or {}
    if not isinstance(raw, dict):
        return {}
    result: dict[str, Scalar] = {}
    for context_field, value in raw.items():
        field_name = str(context_field)
        component = component_for_field(series, field_name)
        if component is not None:
            inferred = _component_dtype(concept_scheme, series, component)
        else:
            inferred = _lookup_concept_dtype(concept_scheme, series, field_name)
        read_as = _effective_read_as({"kind": "constant"}, inferred_dtype=inferred)
        try:
            result[field_name] = coerce_constant(value, read_as=read_as)
        except (ValueError, TypeError) as exc:
            raise ValueError(f"series_context[{field_name!r}]: {exc}") from exc
    return result


def _build_input_record(
    *,
    coordinates: dict[str, Scalar],
    series: dict[str, Any],
    measure_concept: str,
    series_context: dict[str, Scalar],
) -> dict[str, Scalar]:
    record: dict[str, Scalar] = {}
    key_fields = [str(c) for c in (series.get("key") or [])]

    for field_name in key_fields:
        if field_name in coordinates:
            record[field_name] = coordinates[field_name]

    obs_value = coordinates.get(measure_concept)
    if obs_value is not None or measure_concept in coordinates:
        record[measure_concept] = obs_value

    for concept, value in series_context.items():
        record[str(concept)] = value

    structure = series.get("structure") or {}
    for dim in structure.get("dimensions") or []:
        if not isinstance(dim, dict):
            continue
        field_name = effective_dimension_id(dim)
        if _include_in_record(dim, default=True) and field_name in coordinates:
            record[field_name] = coordinates[field_name]

    for attr in structure.get("attributes") or []:
        if not isinstance(attr, dict):
            continue
        field_name = effective_dimension_id(attr)
        if not _include_in_record(attr, default=False):
            continue
        if field_name in coordinates:
            record[field_name] = coordinates[field_name]

    return record


def _build_output_record(
    *,
    coordinates: dict[str, Scalar],
    series: dict[str, Any],
    measure_concept: str,
    series_context: dict[str, Scalar],
) -> dict[str, Scalar]:
    """Build a record with all declared dimensions and attributes (OBS_VALUE filled at runtime)."""
    record: dict[str, Scalar] = {}

    for concept, value in series_context.items():
        record[str(concept)] = value

    structure = series.get("structure") or {}
    for dim in structure.get("dimensions") or []:
        if not isinstance(dim, dict):
            continue
        field_name = effective_dimension_id(dim)
        if field_name in coordinates:
            record[field_name] = coordinates[field_name]

    for attr in structure.get("attributes") or []:
        if not isinstance(attr, dict):
            continue
        field_name = effective_dimension_id(attr)
        if field_name in coordinates:
            record[field_name] = coordinates[field_name]

    if measure_concept not in record:
        record[measure_concept] = coordinates.get(measure_concept)

    return record


def _extract_key(coordinates: dict[str, Scalar], key_fields: list[str]) -> dict[str, Scalar]:
    return {field: coordinates[field] for field in key_fields if field in coordinates}


def resolve_key_domain(
    workbook: Path | str,
    series: dict[str, Any],
    cells: Sequence[str],
    *,
    concept_scheme: dict[str, Any] | None = None,
    graph: DependencyGraph | None = None,
) -> tuple[dict[str, Scalar], ...]:
    """Resolve per-cell key coordinates for `cells` in expansion order.

    Three outcomes, distinguished by type rather than message matching:

    * No key (`key: []`): one empty dict per cell. Expansion order is the
      schedule.
    * Fully resolved: one dict with every declared key field on every cell.
    * Partially resolved: raises `PartialKeyDomainError` naming the series
      and the data cells whose key fields did not resolve.

    Uses the same dimension binds as `resolve_series_binding`, but does not
    intersect with the dependency graph. Structural bind failures
    (`UnknownBindKindError`, missing bind keys) raise immediately.

    Returns:
        Per-cell key dicts in `cells` order. Empty dicts iff no key is
        declared.

    Raises:
        UnknownBindKindError: A key-field bind used an unimplemented kind.
        PartialKeyDomainError: A declared key field is missing on at least
            one cell.
        ValueError: Other key-field bind failures (missing bind keys).
    """
    key_fields = [str(c) for c in (series.get("key") or [])]
    series_id = str(series.get("id") or "")
    if not cells:
        return ()
    if not key_fields:
        return tuple({} for _ in cells)
    structure = series.get("structure") or {}
    points: list[dict[str, Scalar]] = []
    unresolved: dict[str, list[str]] = {}
    with _WorkbookValues(workbook) as reader:
        reader.prefetch(_structure_source_addresses(series, cells), graph=graph)
        series_coordinates: dict[str, Scalar] = {}
        for address in cells:
            coordinates: dict[str, Scalar] = {}
            for dim in structure.get("dimensions") or []:
                if not isinstance(dim, dict):
                    continue
                field_name = effective_dimension_id(dim)
                bind = dim.get("bind")
                if not isinstance(bind, dict):
                    continue
                scope = dim.get("scope")
                if scope == "series" and field_name in series_coordinates:
                    coordinates[field_name] = series_coordinates[field_name]
                    continue
                inferred = _component_dtype(concept_scheme, series, dim)
                try:
                    value = _execute_bind(
                        bind,
                        graph=graph,
                        reader=reader,
                        data_address=address,
                        inferred_read_as=inferred,
                    )
                except UnknownBindKindError as exc:
                    if field_name in key_fields:
                        raise UnknownBindKindError(
                            exc.kind,
                            series_id=series_id,
                            address=address,
                            field_name=field_name,
                        ) from exc
                    continue
                except (KeyError, TypeError) as exc:
                    if field_name in key_fields:
                        raise ValueError(
                            f"series {series_id!r} cell {address}: key field "
                            f"{field_name!r} bind failed: {exc}"
                        ) from exc
                    continue
                except ValueError:
                    continue
                if field_name in key_fields and value is None:
                    continue
                coordinates[field_name] = value
                if scope == "series":
                    series_coordinates[field_name] = value
            point = _extract_key(coordinates, key_fields)
            missing = [field for field in key_fields if field not in point]
            if missing:
                unresolved[address] = missing
            points.append(point)
    if unresolved:
        raise PartialKeyDomainError(series_id, unresolved)
    return tuple(points)


def warn_series_resolution_issues(resolved: SeriesResolution, *, stacklevel: int = 3) -> None:
    """Emit Python warnings for partial-overlap issues on a binding resolution."""
    for issue in resolved["issues"]:
        if issue["level"] != "warning":
            continue
        if issue["code"] in {"partial_graph_overlap", "partial_export_overlap"}:
            warnings.warn(issue["message"], UserWarning, stacklevel=stacklevel)


def _normalize_export_addresses(export_addresses: Iterable[str] | None) -> frozenset[str] | None:
    if export_addresses is None:
        return None
    return frozenset(normalize_address(addr) for addr in export_addresses)


def _select_input_addresses(
    graph: DependencyGraph,
    expanded_addresses: list[str],
    *,
    mode: InputMode,
    validation: dict[str, Any],
    series_id: str,
    candidate_addresses: list[str] | None = None,
) -> tuple[list[str], list[ResolutionIssue]]:
    """Select input binding addresses and report skipped non-leaf cells."""
    issues: list[ResolutionIssue] = []
    pool = expanded_addresses if candidate_addresses is None else candidate_addresses
    if mode == "override":
        selected = [address for address in pool if is_graph_node(graph, address)]
    else:
        intersect = bool(validation.get("intersect_graph_leaves", True))
        if not intersect:
            selected = list(pool)
        else:
            selected = [address for address in pool if is_graph_leaf(graph, address)]
    if mode != "override":
        skipped = len(pool) - len(selected)
        if skipped > 0 and bool(validation.get("warn_on_partial_overlap", True)):
            issues.append(
                make_issue(
                    "warning",
                    "partial_graph_overlap",
                    f"Skipped {skipped} cell(s) in data_range not graph input leaf cells",
                    series_id=series_id,
                )
            )
    return selected, issues


def _select_internal_addresses(
    graph: DependencyGraph,
    expanded_addresses: list[str],
    *,
    validation: dict[str, Any],
    series_id: str,
    candidate_addresses: list[str] | None = None,
) -> tuple[list[str], list[ResolutionIssue]]:
    """Select internal binding addresses and report skipped non-formula cells."""
    issues: list[ResolutionIssue] = []
    pool = expanded_addresses if candidate_addresses is None else candidate_addresses
    intersect = bool(validation.get("intersect_graph_formulas", True))
    if not intersect:
        selected = list(pool)
    else:
        selected = [address for address in pool if is_graph_formula_node(graph, address)]
    skipped = len(pool) - len(selected)
    if skipped > 0 and bool(validation.get("warn_on_partial_overlap", True)):
        issues.append(
            make_issue(
                "warning",
                "partial_graph_overlap",
                f"Skipped {skipped} cell(s) in data_range not graph formula cells",
                series_id=series_id,
            )
        )
    return selected, issues


def _select_addresses(
    graph: DependencyGraph,
    expanded_addresses: list[str],
    *,
    direction: BindingDirection,
    validation: dict[str, Any],
    series_id: str,
    export_addresses: frozenset[str] | None = None,
    input_binding_mode: InputMode = "leaf",
) -> tuple[list[str], list[ResolutionIssue]]:
    issues: list[ResolutionIssue] = []
    if export_addresses is not None:
        if direction in ("input", "constant"):
            exported_addresses = [
                address for address in expanded_addresses if address in export_addresses
            ]
            selected, overlap_issues = _select_input_addresses(
                graph,
                expanded_addresses,
                mode=input_binding_mode if direction == "input" else "leaf",
                validation=validation,
                series_id=series_id,
                candidate_addresses=exported_addresses,
            )
            issues.extend(overlap_issues)
            skipped_export = len(expanded_addresses) - len(exported_addresses)
        else:
            selected = [address for address in expanded_addresses if address in export_addresses]
            skipped_export = len(expanded_addresses) - len(selected)
        if skipped_export > 0 and bool(validation.get("warn_on_partial_overlap", True)):
            issues.append(
                make_issue(
                    "warning",
                    "partial_export_overlap",
                    f"Skipped {skipped_export} cell(s) in data_range not included in codegen export closure",
                    series_id=series_id,
                )
            )
        return selected, issues

    if direction == "input":
        selected, overlap_issues = _select_input_addresses(
            graph,
            expanded_addresses,
            mode=input_binding_mode,
            validation=validation,
            series_id=series_id,
        )
        issues.extend(overlap_issues)
        return selected, issues

    if direction == "constant":
        selected, overlap_issues = _select_input_addresses(
            graph,
            expanded_addresses,
            mode="leaf",
            validation=validation,
            series_id=series_id,
        )
        issues.extend(overlap_issues)
        return selected, issues

    if direction == "internal":
        selected, overlap_issues = _select_internal_addresses(
            graph,
            expanded_addresses,
            validation=validation,
            series_id=series_id,
        )
        issues.extend(overlap_issues)
        return selected, issues

    intersect = bool(validation.get("intersect_graph_nodes", True))
    if not intersect:
        return expanded_addresses, issues
    selected = [address for address in expanded_addresses if is_graph_node(graph, address)]

    skipped = len(expanded_addresses) - len(selected)
    if skipped > 0 and bool(validation.get("warn_on_partial_overlap", True)):
        message = (
            f"Skipped {skipped} cell(s) in data_range not present in graph for {direction} binding"
        )
        issues.append(
            make_issue(
                "warning",
                "partial_graph_overlap",
                message,
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
    export_addresses: Iterable[str] | None = None,
    concept_scheme: dict[str, Any] | None = None,
    reader: _WorkbookValues | None = None,
) -> SeriesResolution:
    """Resolve each participating cell in a series binding to coordinates and record fields.

    Pass `reader` to reuse a shared `_WorkbookValues` across multiple series;
    when omitted, a reader is created for this call and closed before returning.
    """
    series_id = str(series.get("id", ""))
    issues: list[ResolutionIssue] = []
    leaves: list[LeafResolution] = []
    requires_address = False

    if not series_data_ranges(series):
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
        expanded_addresses = apply_series_excludes(
            expand_series_data_ranges_for_graph(graph, series, workbook=workbook),
            series,
        )
    except (ValueError, TypeError) as exc:
        return {
            "series_id": series_id,
            "ok": False,
            "requires_address": True,
            "leaves": [],
            "issues": [make_issue("error", "invalid_data_range", str(exc), series_id=series_id)],
        }

    validation = effective_validation(series)
    export_set = _normalize_export_addresses(export_addresses)
    addresses, overlap_issues = _select_addresses(
        graph,
        expanded_addresses,
        direction=direction,
        validation=validation,
        series_id=series_id,
        export_addresses=export_set,
        input_binding_mode=input_mode(series) if direction == "input" else "leaf",
    )
    issues.extend(overlap_issues)
    if not addresses:
        intersection_label = (
            "codegen export intersection" if export_set is not None else "graph intersection"
        )
        issues.append(
            make_issue(
                "warning",
                "no_resolved_cells",
                f"No resolved {direction} cells in data_range after {intersection_label}",
                series_id=series_id,
            )
        )

    structure = series.get("structure") or {}
    measure = structure.get("measure") or {}
    measure_concept = str(measure.get("concept", "OBS_VALUE"))
    measure_bind = measure.get("bind") or {"kind": "data_cell"}
    measure_dtype = measure.get("dtype")
    measure_inferred_read = str(measure_dtype) if measure_dtype is not None else None
    key_fields = [str(c) for c in (series.get("key") or [])]
    require_unique_key = bool(validation.get("require_unique_key", True))

    owns_reader = reader is None
    if owns_reader:
        reader = _WorkbookValues(workbook)
    active_reader = reader

    try:
        series_coordinates: dict[str, Scalar] = {}
        seen_keys: dict[tuple[tuple[str, Scalar], ...], str] = {}

        try:
            coerced_series_context = _coerce_series_context(series, concept_scheme=concept_scheme)
        except ValueError as exc:
            return {
                "series_id": series_id,
                "ok": False,
                "requires_address": True,
                "leaves": [],
                "issues": [
                    make_issue(
                        "error",
                        "series_context_coercion_failed",
                        str(exc),
                        series_id=series_id,
                    )
                ],
            }

        build_record = _build_input_record if direction == "input" else _build_output_record

        bind_failures: dict[str, list[str]] = {}
        active_reader.prefetch(
            _structure_source_addresses(
                series,
                addresses,
                include_measure=direction == "input",
                include_attributes=True,
            ),
            graph=graph,
        )
        for address in addresses:
            coordinates: dict[str, Scalar] = {}
            try:
                if direction == "input":
                    coordinates[measure_concept] = _execute_bind(
                        measure_bind if isinstance(measure_bind, dict) else {"kind": "data_cell"},
                        graph=graph,
                        reader=active_reader,
                        data_address=address,
                        inferred_read_as=measure_inferred_read,
                    )

                for dim in structure.get("dimensions") or []:
                    if not isinstance(dim, dict):
                        continue
                    field_name = effective_dimension_id(dim)
                    bind = dim.get("bind")
                    if not isinstance(bind, dict):
                        continue
                    scope = dim.get("scope")
                    if scope == "series" and field_name in series_coordinates:
                        coordinates[field_name] = series_coordinates[field_name]
                        continue
                    inferred = _component_dtype(concept_scheme, series, dim)
                    value = _execute_bind(
                        bind,
                        graph=graph,
                        reader=active_reader,
                        data_address=address,
                        inferred_read_as=inferred,
                    )
                    coordinates[field_name] = value
                    if scope == "series":
                        series_coordinates[field_name] = value

                for attr in structure.get("attributes") or []:
                    if not isinstance(attr, dict):
                        continue
                    field_name = effective_dimension_id(attr)
                    bind = _attribute_bind(attr)
                    if bind is None:
                        continue
                    inferred = _component_dtype(concept_scheme, series, attr)
                    coordinates[field_name] = _execute_bind(
                        bind,
                        graph=graph,
                        reader=active_reader,
                        data_address=address,
                        inferred_read_as=inferred,
                    )
            except (KeyError, ValueError, TypeError) as exc:
                bind_failures.setdefault(str(exc), []).append(address)
                continue

            key = _extract_key(coordinates, key_fields)
            record = build_record(
                coordinates=coordinates,
                series=series,
                measure_concept=measure_concept,
                series_context=coerced_series_context,
            )
            leaves.append(
                {
                    "address": address,
                    "coordinates": coordinates,
                    "key": key,
                    "record": record,
                }
            )

            if require_unique_key and key_fields:
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

        issues.extend(
            _bind_resolution_issues(bind_failures, series_id=series_id),
        )

        layout = series.get("layout")
        if layout == "scalar" and not key_fields and len(leaves) > 1:
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
    finally:
        if owns_reader:
            active_reader.close()


def _series_supports_direction(series: dict[str, Any], direction: BindingDirection) -> bool:
    if direction == "input":
        return has_input_direction(series)
    if direction == "internal":
        return has_internal_direction(series)
    if direction == "constant":
        return has_constant_direction(series)
    return has_output_direction(series)


def resolve_series_bindings(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str | None = None,
    direction: BindingDirection = "input",
    export_addresses: Iterable[str] | None = None,
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
    concept_scheme = bindings.get("concept_scheme")
    if not isinstance(concept_scheme, dict):
        concept_scheme = None
    with _WorkbookValues(workbook) as reader:
        for series in bindings.get("series", []):
            if not isinstance(series, dict):
                continue
            if not _series_supports_direction(series, direction):
                continue
            result = resolve_series_binding(
                graph,
                workbook,
                series,
                direction=direction,
                export_addresses=export_addresses,
                concept_scheme=concept_scheme,
                reader=reader,
            )
            series_results.append(result)
            all_issues.extend(result["issues"])

    ok = all(r["ok"] for r in series_results) and not any(i["level"] == "error" for i in all_issues)
    return {"ok": ok, "series": series_results, "issues": all_issues}
