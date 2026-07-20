"""Emit generated input setters/readers and output compute functions from series bindings."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import TYPE_CHECKING, cast

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.compute_codegen import emit_computes_block
from excel_grapher.series_bindings.groups import bindings_export_order
from excel_grapher.series_bindings.normalize import has_input_direction, has_output_direction
from excel_grapher.series_bindings.setter_codegen import emit_setters_block
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

if TYPE_CHECKING:
    from excel_grapher.series_bindings.docstring_renderers import SeriesDocstringRendererSpec
    from excel_grapher.series_bindings.docstrings import SeriesBindingDocstringCallbackSpec
    from excel_grapher.series_bindings.output_helper_index import (
        OutputHelperIndex,
        OutputHelperSpec,
    )


def emit_series_bindings_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
    include_helpers: bool = True,
    include_readers: bool = True,
    include_leaf_indexes: bool = True,
    include_leaves_tables: bool = True,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
    helper_index: OutputHelperIndex | None = None,
    address_helpers: Mapping[str, OutputHelperSpec] | None = None,
) -> list[str]:
    """Emit setter, reader, and/or output compute functions for a binding manifest.

    When `include_helpers` is true the coercion helpers and type aliases are inlined
    (single-file export). When false only the public setter/reader/compute functions are
    emitted; the helpers and aliases are expected to be importable from a separate
    module (the multi-module export's `_api_helpers`).

    When `include_readers` / `include_leaf_indexes` are false, those symbols are omitted
    so a dedicated `_readers` module can own them (modular export).

    When `include_leaves_tables` is false, `_OUTPUT_LEAVES_*` tables are omitted so a
    dedicated `_output_leaves` module can own them (modular export).
    """
    series_list = bindings_export_order(bindings)
    emit_input = any(has_input_direction(s) for s in series_list)
    emit_output = any(has_output_direction(s) for s in series_list)
    if not emit_input and not emit_output:
        return []
    bindings = cast(WorkbookSeriesBindings, {**bindings, "series": series_list})

    lines: list[str] = []
    include_aliases = include_helpers
    if emit_input:
        lines.extend(
            emit_setters_block(
                graph,
                workbook,
                bindings,
                export_addresses=export_addresses,
                include_type_aliases=include_aliases,
                include_helpers=include_helpers,
                include_readers=include_readers,
                include_leaf_indexes=include_leaf_indexes,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
            )
        )
        include_aliases = False
    if emit_output:
        lines.extend(
            emit_computes_block(
                graph,
                workbook,
                bindings,
                export_addresses=export_addresses,
                include_type_aliases=include_aliases,
                include_leaves_tables=include_leaves_tables,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
                helper_index=helper_index,
                address_helpers=address_helpers,
            )
        )
    return lines
