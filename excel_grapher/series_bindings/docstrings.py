"""Structured docstring callbacks for generated series-binding API functions."""

from __future__ import annotations

import textwrap
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal, Protocol, TypeAlias, runtime_checkable

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import effective_dimension_id
from excel_grapher.series_bindings.types import SeriesResolution, WorkbookSeriesBindings

if TYPE_CHECKING:
    from excel_grapher.series_bindings.docstring_renderers import SeriesDocstringRendererSpec

SeriesFunctionKind = Literal["setter", "compute"]


@dataclass(frozen=True, slots=True)
class FieldContract:
    concept_name: str
    dtype: str | None
    required: bool
    expected_value: object | None


@dataclass(frozen=True, slots=True)
class SeriesBindingDocstringContract:
    series_id: str
    function_name: str
    function_kind: SeriesFunctionKind
    data_range: str
    layout: str
    value_type: str | None
    required_fields: tuple[str, ...]
    fields: Mapping[str, FieldContract]
    example_records: tuple[Mapping[str, object], ...]
    notes: str


@dataclass(frozen=True, slots=True)
class FieldDoc:
    description: str


@dataclass(frozen=True, slots=True)
class SeriesFunctionDoc:
    summary: str
    purpose: str
    record_matching: str
    field_descriptions: Mapping[str, FieldDoc] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class SeriesBindingDocstringContext:
    graph: DependencyGraph
    workbook: Path | str
    bindings: WorkbookSeriesBindings
    series: dict[str, Any]
    resolution: SeriesResolution
    contract: SeriesBindingDocstringContract
    function_kind: SeriesFunctionKind
    function_name: str


@runtime_checkable
class SeriesBindingDocstringCallback(Protocol):
    def __call__(self, ctx: SeriesBindingDocstringContext) -> SeriesFunctionDoc | None: ...


SeriesBindingDocstringCallbackSpec: TypeAlias = str | SeriesBindingDocstringCallback

_callbacks: dict[str, SeriesBindingDocstringCallback] = {}


def register_series_docstring_callback(
    name: str,
    callback: SeriesBindingDocstringCallback,
    *,
    replace: bool = False,
) -> None:
    if name in _callbacks and not replace:
        raise ValueError(f"duplicate series docstring callback: {name!r}")
    _callbacks[name] = callback


def list_series_docstring_callbacks() -> tuple[str, ...]:
    return tuple(sorted(_callbacks))


def resolve_series_docstring_callback(
    callback: SeriesBindingDocstringCallbackSpec,
) -> SeriesBindingDocstringCallback:
    """Look up a registered callback name or return a direct callback object."""
    if isinstance(callback, str):
        if callback not in _callbacks:
            known = ", ".join(sorted(_callbacks))
            raise ValueError(f"Unknown series docstring callback: {callback!r}. Known: {known}")
        return _callbacks[callback]
    return callback


def run_series_docstring_callback(
    name: str,
    ctx: SeriesBindingDocstringContext,
) -> SeriesFunctionDoc | None:
    """Execute a registered series docstring callback by name."""
    return resolve_series_docstring_callback(name)(ctx)


def unregister_series_docstring_callback(name: str) -> None:
    """Remove a registered callback (for tests and notebook cleanup)."""
    _callbacks.pop(name, None)


def _unique(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


def _concept_lookup(bindings: WorkbookSeriesBindings) -> dict[str, dict[str, Any]]:
    scheme = bindings.get("concept_scheme") or {}
    concepts = scheme.get("concepts") or []
    return {
        str(concept["id"]): concept
        for concept in concepts
        if isinstance(concept, dict) and "id" in concept
    }


def _measure_concept(series: dict[str, Any]) -> str:
    measure = (series.get("structure") or {}).get("measure") or {}
    return str(measure.get("concept") or "OBS_VALUE")


def _expected_record_values(series: dict[str, Any]) -> dict[str, object]:
    expected = dict(series.get("series_context") or {})
    for attribute in (series.get("structure") or {}).get("attributes") or []:
        if isinstance(attribute, dict) and "value" in attribute:
            expected[effective_dimension_id(attribute)] = attribute["value"]
    return expected


def _record_field_names(series: dict[str, Any]) -> list[str]:
    measure = _measure_concept(series)
    dimensions = [
        effective_dimension_id(dimension)
        for dimension in (series.get("structure") or {}).get("dimensions") or []
        if isinstance(dimension, dict) and "concept" in dimension
    ]
    attributes = [
        effective_dimension_id(attribute)
        for attribute in (series.get("structure") or {}).get("attributes") or []
        if isinstance(attribute, dict) and "concept" in attribute
    ]
    required = [*(series.get("key") or []), measure]
    return _unique([*required, *dimensions, *attributes])


def _concept_id_for_field(series: dict[str, Any], field_name: str) -> str:
    """Map a record field name to the concept id used for scheme lookups."""
    structure = series.get("structure") or {}
    components = [
        *(structure.get("dimensions") or []),
        *(structure.get("attributes") or []),
    ]
    for component in components:
        if not isinstance(component, dict):
            continue
        if effective_dimension_id(component) != field_name:
            continue
        concept = component.get("concept")
        if concept:
            return str(concept)
    return field_name


def _dimension_bind_read(series: dict[str, Any], field_name: str) -> str | None:
    for dimension in (series.get("structure") or {}).get("dimensions") or []:
        if not isinstance(dimension, dict):
            continue
        if effective_dimension_id(dimension) != field_name:
            continue
        bind = dimension.get("bind")
        if isinstance(bind, dict) and bind.get("read") is not None:
            return str(bind["read"])
    return None


def _field_dtype(
    series: dict[str, Any],
    concepts: dict[str, dict[str, Any]],
    field_name: str,
) -> str | None:
    measure = (series.get("structure") or {}).get("measure") or {}
    if field_name == _measure_concept(series):
        measure_dtype = measure.get("dtype")
        if measure_dtype is not None:
            return str(measure_dtype)
        concept = concepts.get(field_name)
        if concept is not None and concept.get("dtype") is not None:
            return str(concept["dtype"])
        return None
    concept = concepts.get(_concept_id_for_field(series, field_name))
    if concept is not None and concept.get("dtype") is not None:
        return str(concept["dtype"])
    bind_read = _dimension_bind_read(series, field_name)
    if bind_read in {"string", "int", "float", "number", "bool", "datetime"}:
        return bind_read
    return None


def _example_records_from_resolution(
    resolution: SeriesResolution,
    required_fields: tuple[str, ...],
    *,
    max_records: int = 2,
) -> tuple[dict[str, object], ...]:
    examples: list[dict[str, object]] = []
    for leaf in resolution["leaves"][:max_records]:
        record = leaf["record"]
        examples.append({field_name: record[field_name] for field_name in required_fields})
    return tuple(examples)


def derive_doc_contract(
    series: dict[str, Any],
    *,
    function_kind: SeriesFunctionKind,
    function_name: str,
    resolution: SeriesResolution,
    bindings: WorkbookSeriesBindings,
) -> SeriesBindingDocstringContract:
    concepts = _concept_lookup(bindings)
    measure = _measure_concept(series)
    required_fields = tuple(str(field) for field in (series.get("key") or [])) + (measure,)
    expected_values = _expected_record_values(series)
    field_contracts: dict[str, FieldContract] = {}
    for field_name in _record_field_names(series):
        concept = concepts.get(_concept_id_for_field(series, field_name))
        concept_name = (
            str(concept["name"]) if concept is not None and concept.get("name") else field_name
        )
        field_contracts[field_name] = FieldContract(
            concept_name=concept_name,
            dtype=_field_dtype(series, concepts, field_name),
            required=field_name in required_fields,
            expected_value=expected_values.get(field_name),
        )

    return SeriesBindingDocstringContract(
        series_id=str(series.get("id", "")),
        function_name=function_name,
        function_kind=function_kind,
        data_range=str(series.get("data_range", "")),
        layout=str(series.get("layout", "")),
        value_type=_field_dtype(series, concepts, measure),
        required_fields=required_fields,
        fields=field_contracts,
        example_records=_example_records_from_resolution(resolution, required_fields),
        notes=str(series.get("notes") or series.get("sdmx_notes") or ""),
    )


def emit_docstring_literal(doc: str) -> list[str]:
    escaped = doc.replace('"""', '\\"""')
    if "\n" not in escaped:
        return [f'    """{escaped}"""']
    body = textwrap.indent(escaped, "    ")
    return [f'    """{body.lstrip()}', '    """']


def _default_docstring(
    series: dict[str, Any],
    *,
    function_kind: SeriesFunctionKind,
    series_id: str,
) -> str:
    notes = series.get("notes") or series.get("sdmx_notes")
    if notes:
        return str(notes)
    if function_kind == "setter":
        return f"Apply records for {series_id}."
    return f"Compute records for {series_id}."


def _series_notes_callback(ctx: SeriesBindingDocstringContext) -> SeriesFunctionDoc | None:
    notes = ctx.contract.notes
    if not notes:
        return None
    return SeriesFunctionDoc(
        summary=notes,
        purpose="",
        record_matching="",
        field_descriptions={},
    )


def resolve_series_function_docstring(
    *,
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    series: dict[str, Any],
    resolution: SeriesResolution,
    function_kind: SeriesFunctionKind,
    function_name: str,
    callback_spec: SeriesBindingDocstringCallbackSpec | None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> str | None:
    if callback_spec is None:
        return _default_docstring(
            series,
            function_kind=function_kind,
            series_id=str(resolution["series_id"]),
        )

    from excel_grapher.series_bindings.docstring_renderers import (
        resolve_series_docstring_renderer,
    )

    contract = derive_doc_contract(
        series,
        function_kind=function_kind,
        function_name=function_name,
        resolution=resolution,
        bindings=bindings,
    )
    ctx = SeriesBindingDocstringContext(
        graph=graph,
        workbook=workbook,
        bindings=bindings,
        series=series,
        resolution=resolution,
        contract=contract,
        function_kind=function_kind,
        function_name=function_name,
    )
    structured = resolve_series_docstring_callback(callback_spec)(ctx)
    if structured is None:
        return None
    renderer = resolve_series_docstring_renderer(docstring_renderer)
    return renderer.render(contract, structured, series=series)


def _register_builtin_callbacks() -> None:
    register_series_docstring_callback("series_notes", _series_notes_callback)


_register_builtin_callbacks()

__all__ = [
    "FieldContract",
    "FieldDoc",
    "SeriesBindingDocstringCallback",
    "SeriesBindingDocstringCallbackSpec",
    "SeriesBindingDocstringContext",
    "SeriesBindingDocstringContract",
    "SeriesFunctionDoc",
    "SeriesFunctionKind",
    "derive_doc_contract",
    "emit_docstring_literal",
    "list_series_docstring_callbacks",
    "register_series_docstring_callback",
    "resolve_series_docstring_callback",
    "resolve_series_function_docstring",
    "run_series_docstring_callback",
    "unregister_series_docstring_callback",
]
