"""Series binding manifest: load, validate, and canonicalize workbook setter specs."""

from __future__ import annotations

from excel_grapher.series_bindings.bindings_codegen import emit_series_bindings_block
from excel_grapher.series_bindings.canonical import bindings_canonical_sha256
from excel_grapher.series_bindings.compute_codegen import (
    emit_compute_function,
    emit_computes_block,
    generate_computes_module,
)
from excel_grapher.series_bindings.docstring_renderers import (
    GoogleSeriesDocstringRenderer,
    NumpySeriesDocstringRenderer,
    PlainSeriesDocstringRenderer,
    RstSeriesDocstringRenderer,
    SeriesDocstringRenderCallable,
    SeriesDocstringRenderer,
    SeriesDocstringRendererName,
    SeriesDocstringRendererSpec,
    resolve_series_docstring_renderer,
)
from excel_grapher.series_bindings.docstrings import (
    FieldDoc,
    SeriesBindingDocstringContext,
    SeriesBindingDocstringContract,
    SeriesFunctionDoc,
    list_series_docstring_callbacks,
    register_series_docstring_callback,
)
from excel_grapher.series_bindings.input_series import derive_input_series
from excel_grapher.series_bindings.load import (
    SeriesBindingsLoadError,
    load_series_bindings,
    merge_series_binding_documents,
    parse_bindings_file,
)
from excel_grapher.series_bindings.normalize import (
    has_input_direction,
    has_output_direction,
    merge_series_entries,
    normalize_bindings_document,
    normalize_series_entry,
)
from excel_grapher.series_bindings.output_series import derive_output_series
from excel_grapher.series_bindings.ranges import expand_data_range, expand_data_range_for_graph
from excel_grapher.series_bindings.records_types import Record, Records, Scalar
from excel_grapher.series_bindings.resolve import resolve_series_binding, resolve_series_bindings
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    format_schema_errors,
    validate_bindings_document,
)
from excel_grapher.series_bindings.setter_codegen import (
    emit_setter_function,
    emit_setters_block,
    generate_setters_module,
)
from excel_grapher.series_bindings.types import (
    InputSeries,
    InputSeriesCell,
    LeafResolution,
    OutputSeries,
    OutputSeriesCell,
    ResolutionIssue,
    ResolutionReport,
    SeriesResolution,
    ValidationIssue,
    ValidationLevel,
    ValidationReport,
    WorkbookSeriesBindings,
)
from excel_grapher.series_bindings.validate import validate_series_bindings
from excel_grapher.series_bindings.versions import (
    IMPLEMENTED_BIND_KINDS,
    IMPLEMENTED_LAYOUTS,
    PLANNED_BIND_KINDS,
    PLANNED_LAYOUTS,
    SUPPORTED_SCHEMA_VERSIONS,
    is_bind_implemented,
    is_layout_implemented,
)

__all__ = [
    "InputSeries",
    "InputSeriesCell",
    "OutputSeries",
    "OutputSeriesCell",
    "Record",
    "Records",
    "IMPLEMENTED_BIND_KINDS",
    "IMPLEMENTED_LAYOUTS",
    "PLANNED_BIND_KINDS",
    "PLANNED_LAYOUTS",
    "SUPPORTED_SCHEMA_VERSIONS",
    "FieldDoc",
    "GoogleSeriesDocstringRenderer",
    "NumpySeriesDocstringRenderer",
    "PlainSeriesDocstringRenderer",
    "RstSeriesDocstringRenderer",
    "SeriesDocstringRenderCallable",
    "SeriesDocstringRenderer",
    "SeriesDocstringRendererName",
    "SeriesDocstringRendererSpec",
    "SeriesBindingDocstringContext",
    "SeriesBindingDocstringContract",
    "SeriesFunctionDoc",
    "LeafResolution",
    "ResolutionIssue",
    "ResolutionReport",
    "Scalar",
    "SeriesResolution",
    "SeriesBindingsLoadError",
    "SeriesBindingsSchemaError",
    "ValidationIssue",
    "ValidationLevel",
    "ValidationReport",
    "WorkbookSeriesBindings",
    "bindings_canonical_sha256",
    "derive_input_series",
    "derive_output_series",
    "emit_compute_function",
    "emit_computes_block",
    "emit_series_bindings_block",
    "emit_setter_function",
    "emit_setters_block",
    "generate_computes_module",
    "has_input_direction",
    "has_output_direction",
    "merge_series_entries",
    "normalize_bindings_document",
    "normalize_series_entry",
    "expand_data_range",
    "expand_data_range_for_graph",
    "format_schema_errors",
    "generate_setters_module",
    "list_series_docstring_callbacks",
    "load_series_bindings",
    "merge_series_binding_documents",
    "parse_bindings_file",
    "register_series_docstring_callback",
    "resolve_series_binding",
    "resolve_series_bindings",
    "resolve_series_docstring_renderer",
    "validate_bindings_document",
    "validate_series_bindings",
    "is_bind_implemented",
    "is_layout_implemented",
]
