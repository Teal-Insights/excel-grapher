"""Issue 842 Tier 2: unused library symbols must stay gone."""

from __future__ import annotations

import excel_grapher.core.addressing as addressing
import excel_grapher.core.cell_types as cell_types
import excel_grapher.exporter.export_runtime.errors as export_errors
import excel_grapher.exporter.export_runtime.values as export_values
import excel_grapher.exporter.inverted_tree.emit as emit
import excel_grapher.grapher.builder as builder
import excel_grapher.grapher.compression as compression
import excel_grapher.grapher.dynamic_refs as dynamic_refs
import excel_grapher.grapher.parser as parser
import excel_grapher.grapher.resolver as resolver
import excel_grapher.grapher.type_analysis_cache as type_analysis_cache
import excel_grapher.series_bindings.output_helper_index as output_helper_index
import excel_grapher.series_bindings.workflow as workflow


def test_tier2_unused_symbols_are_removed() -> None:
    assert not hasattr(dynamic_refs, "_infer_choose_numeric_domain")
    assert not hasattr(workflow, "output_binding_covered_addresses")
    assert not hasattr(emit, "_HOLE_DOC_LABELS")
    assert not hasattr(export_errors, "raise_if_sentinel")
    assert hasattr(export_errors, "raise_if_sentinel_float")
    assert not hasattr(export_values, "_convergence_delta")
    assert not hasattr(resolver, "qualify_cell_ref")
    assert not hasattr(compression.OptimalCompressionRecord, "ensure_snapshot")
    assert not hasattr(parser, "_QUOTED_SHEET")
    assert not hasattr(parser, "_needs_quoting")
    assert not hasattr(parser, "_format_ref")
    assert not hasattr(builder, "_logger")
    assert not hasattr(builder, "_VOLATILE_DYNAMIC_REF_FUNCS")
    assert not hasattr(cell_types, "_normalize_cell_address")
    assert not hasattr(addressing, "_split_sheet_qualified_address")
    assert not hasattr(output_helper_index, "OutputHelperFallbackReason")
    assert "disabled" not in type_analysis_cache.CacheStats.__dataclass_fields__
