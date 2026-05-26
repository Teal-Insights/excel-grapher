"""Series binding manifest: load, validate, and canonicalize workbook setter specs."""

from __future__ import annotations

from excel_grapher.series_bindings.canonical import bindings_canonical_sha256
from excel_grapher.series_bindings.input_series import derive_input_series
from excel_grapher.series_bindings.load import (
    SeriesBindingsLoadError,
    load_series_bindings,
    merge_series_binding_documents,
    parse_bindings_file,
)
from excel_grapher.series_bindings.ranges import expand_data_range, expand_data_range_for_graph
from excel_grapher.series_bindings.resolve import resolve_series_binding, resolve_series_bindings
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    format_schema_errors,
    validate_bindings_document,
)
from excel_grapher.series_bindings.setter_codegen import (
    Records,
    emit_setter_function,
    emit_setters_block,
    generate_setters_module,
)
from excel_grapher.series_bindings.types import (
    InputSeries,
    InputSeriesCell,
    LeafResolution,
    ResolutionIssue,
    ResolutionReport,
    Scalar,
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
    "IMPLEMENTED_BIND_KINDS",
    "IMPLEMENTED_LAYOUTS",
    "PLANNED_BIND_KINDS",
    "PLANNED_LAYOUTS",
    "SUPPORTED_SCHEMA_VERSIONS",
    "LeafResolution",
    "Records",
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
    "emit_setter_function",
    "emit_setters_block",
    "expand_data_range",
    "expand_data_range_for_graph",
    "format_schema_errors",
    "generate_setters_module",
    "load_series_bindings",
    "merge_series_binding_documents",
    "parse_bindings_file",
    "resolve_series_binding",
    "resolve_series_bindings",
    "validate_bindings_document",
    "validate_series_bindings",
    "is_bind_implemented",
    "is_layout_implemented",
]
