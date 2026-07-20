"""Series binding manifest: load, validate, and canonicalize workbook setter specs."""

from __future__ import annotations

from excel_grapher.series_bindings.bindings_codegen import emit_series_bindings_block
from excel_grapher.series_bindings.canonical import bindings_canonical_sha256
from excel_grapher.series_bindings.compute_codegen import (
    emit_compute_function,
    emit_computes_block,
    emit_output_leaves_block,
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
    SeriesBindingDocstringCallback,
    SeriesBindingDocstringCallbackSpec,
    SeriesBindingDocstringContext,
    SeriesBindingDocstringContract,
    SeriesFunctionDoc,
    list_series_docstring_callbacks,
    register_series_docstring_callback,
    resolve_series_docstring_callback,
    unregister_series_docstring_callback,
)
from excel_grapher.series_bindings.groups import (
    GroupMember,
    GroupNode,
    GroupsManifest,
    bindings_export_order,
    bindings_have_groups,
    group_manifest,
    group_slug,
    grouped_public_names,
)
from excel_grapher.series_bindings.input_coerce import coerce_setter_input
from excel_grapher.series_bindings.input_series import derive_input_series
from excel_grapher.series_bindings.internal_series import derive_internal_series
from excel_grapher.series_bindings.load import (
    SeriesBindingsLoadError,
    load_series_bindings,
    merge_series_binding_documents,
    parse_bindings_file,
)
from excel_grapher.series_bindings.normalize import (
    has_input_direction,
    has_internal_direction,
    has_output_direction,
    merge_series_entries,
    normalize_bindings_document,
    normalize_series_entry,
)
from excel_grapher.series_bindings.output_helper_index import (
    OutputHelperCallResolution,
    OutputHelperIndex,
    OutputHelperLeafEntry,
    OutputHelperSpec,
    build_output_helper_index,
    format_output_helper_call_form,
    helper_spec_from_series,
    output_helper_names,
    resolve_output_helper_ref,
)
from excel_grapher.series_bindings.output_series import derive_output_series
from excel_grapher.series_bindings.ranges import expand_data_range, expand_data_range_for_graph
from excel_grapher.series_bindings.reader_index import (
    ReaderCallResolution,
    ReaderIndex,
    ReaderLeafEntry,
    ReaderRangeEntry,
    build_reader_index,
    format_reader_call_form,
    resolve_reader_ref,
)
from excel_grapher.series_bindings.resolve import resolve_series_binding, resolve_series_bindings
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    format_schema_errors,
    validate_bindings_document,
)
from excel_grapher.series_bindings.setter_codegen import (
    emit_reader_function,
    emit_reader_range_function,
    emit_setter_function,
    emit_setter_helpers,
    emit_setters_block,
    generate_setters_module,
)
from excel_grapher.series_bindings.setter_input_types import Layout, SeriesInput
from excel_grapher.series_bindings.types import (
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

_LAZY_WORKFLOW_EXPORTS = {
    "BindingsCheckResult",
    "run_binding_checks",
    "validate_bindings_workbook",
}


def __getattr__(name: str):
    if name in _LAZY_WORKFLOW_EXPORTS:
        from excel_grapher.series_bindings import workflow as workflow_module

        return getattr(workflow_module, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = [
    "InputSeries",
    "InputSeriesCell",
    "InternalSeries",
    "InternalSeriesCell",
    "OutputSeries",
    "OutputSeriesCell",
    "OutputHelperCallResolution",
    "OutputHelperIndex",
    "OutputHelperLeafEntry",
    "OutputHelperSpec",
    "ReaderCallResolution",
    "ReaderIndex",
    "ReaderLeafEntry",
    "ReaderRangeEntry",
    "Record",
    "Records",
    "IMPLEMENTED_BIND_KINDS",
    "IMPLEMENTED_LAYOUTS",
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
    "SeriesBindingDocstringCallback",
    "SeriesBindingDocstringCallbackSpec",
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
    "BindingsCheckResult",
    "GroupMember",
    "GroupNode",
    "GroupsManifest",
    "bindings_canonical_sha256",
    "bindings_export_order",
    "bindings_have_groups",
    "build_output_helper_index",
    "build_reader_index",
    "group_manifest",
    "group_slug",
    "grouped_public_names",
    "Layout",
    "SeriesInput",
    "coerce_setter_input",
    "derive_input_series",
    "derive_internal_series",
    "derive_output_series",
    "format_output_helper_call_form",
    "format_reader_call_form",
    "emit_compute_function",
    "emit_computes_block",
    "emit_output_leaves_block",
    "emit_series_bindings_block",
    "emit_reader_function",
    "emit_reader_range_function",
    "emit_setter_function",
    "emit_setter_helpers",
    "emit_setters_block",
    "generate_computes_module",
    "has_input_direction",
    "has_internal_direction",
    "has_output_direction",
    "helper_spec_from_series",
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
    "output_helper_names",
    "parse_bindings_file",
    "register_series_docstring_callback",
    "resolve_series_docstring_callback",
    "unregister_series_docstring_callback",
    "resolve_output_helper_ref",
    "resolve_reader_ref",
    "resolve_series_binding",
    "resolve_series_bindings",
    "resolve_series_docstring_renderer",
    "run_binding_checks",
    "validate_bindings_document",
    "validate_bindings_workbook",
    "validate_series_bindings",
    "is_bind_implemented",
    "is_layout_implemented",
]
