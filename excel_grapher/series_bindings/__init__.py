"""Series binding manifest: load, validate, and canonicalize workbook series specs."""

from __future__ import annotations

from excel_grapher.series_bindings.canonical import bindings_canonical_sha256
from excel_grapher.series_bindings.constant_series import derive_constant_series
from excel_grapher.series_bindings.groups import (
    GroupMember,
    GroupNode,
    GroupsManifest,
    bindings_export_order,
    bindings_have_groups,
    group_manifest,
    group_slug,
)
from excel_grapher.series_bindings.input_coerce import Layout, SeriesInput, coerce_setter_input
from excel_grapher.series_bindings.input_series import derive_input_series
from excel_grapher.series_bindings.internal_series import derive_internal_series
from excel_grapher.series_bindings.load import (
    SeriesBindingsLoadError,
    load_series_bindings,
    merge_series_binding_documents,
    parse_bindings_file,
)
from excel_grapher.series_bindings.normalize import (
    has_constant_direction,
    has_input_direction,
    has_internal_direction,
    has_output_direction,
    merge_series_entries,
    normalize_bindings_document,
    normalize_series_entry,
)
from excel_grapher.series_bindings.output_series import derive_output_series
from excel_grapher.series_bindings.ranges import (
    expand_data_range,
    expand_data_range_for_graph,
    series_data_ranges,
)
from excel_grapher.series_bindings.resolve import resolve_series_binding, resolve_series_bindings
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    format_schema_errors,
    validate_bindings_document,
)
from excel_grapher.series_bindings.types import (
    ConstantSeries,
    ConstantSeriesCell,
    InputSeries,
    InputSeriesCell,
    InternalSeries,
    InternalSeriesCell,
    LeafResolution,
    OutputSeries,
    OutputSeriesCell,
    Record,
    Records,
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
    SUPPORTED_SCHEMA_VERSIONS,
    is_bind_implemented,
    is_layout_implemented,
)
from excel_grapher.series_bindings.workflow import (
    BindingsCheckResult,
    run_binding_checks,
    series_binding_public_addresses,
    validate_bindings_workbook,
)

__all__ = [
    "ConstantSeries",
    "ConstantSeriesCell",
    "InputSeries",
    "InputSeriesCell",
    "InternalSeries",
    "InternalSeriesCell",
    "OutputSeries",
    "OutputSeriesCell",
    "Record",
    "Records",
    "IMPLEMENTED_BIND_KINDS",
    "IMPLEMENTED_LAYOUTS",
    "SUPPORTED_SCHEMA_VERSIONS",
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
    "BindingsCheckResult",
    "GroupMember",
    "GroupNode",
    "GroupsManifest",
    "bindings_canonical_sha256",
    "bindings_export_order",
    "bindings_have_groups",
    "group_manifest",
    "group_slug",
    "Layout",
    "SeriesInput",
    "coerce_setter_input",
    "derive_constant_series",
    "derive_input_series",
    "derive_internal_series",
    "derive_output_series",
    "has_constant_direction",
    "has_input_direction",
    "has_internal_direction",
    "has_output_direction",
    "merge_series_entries",
    "normalize_bindings_document",
    "normalize_series_entry",
    "expand_data_range",
    "expand_data_range_for_graph",
    "series_data_ranges",
    "format_schema_errors",
    "load_series_bindings",
    "merge_series_binding_documents",
    "parse_bindings_file",
    "resolve_series_binding",
    "resolve_series_bindings",
    "run_binding_checks",
    "series_binding_public_addresses",
    "validate_bindings_document",
    "validate_bindings_workbook",
    "validate_series_bindings",
    "is_bind_implemented",
    "is_layout_implemented",
]
