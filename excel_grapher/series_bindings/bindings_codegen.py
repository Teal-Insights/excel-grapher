"""Emit generated input setters/readers and output compute functions from series bindings."""

from __future__ import annotations

from collections.abc import Iterable
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


def emit_series_bindings_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
    include_helpers: bool = True,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit setter, reader, and/or output compute functions for a binding manifest.

    When `include_helpers` is true the coercion helpers and type aliases are inlined
    (single-file export). When false only the public setter/reader/compute functions are
    emitted; the helpers and aliases are expected to be importable from a separate
    module (the multi-module export's `_api_helpers`).
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
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
            )
        )
    return lines
