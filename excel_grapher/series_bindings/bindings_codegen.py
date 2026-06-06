"""Emit generated input setters and output compute functions from series bindings."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
from typing import TYPE_CHECKING

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.compute_codegen import emit_computes_block
from excel_grapher.series_bindings.normalize import has_input_direction, has_output_direction
from excel_grapher.series_bindings.setter_codegen import emit_setters_block
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

if TYPE_CHECKING:
    from excel_grapher.series_bindings.docstring_renderers import SeriesDocstringRendererSpec


def emit_series_bindings_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
    series_docstring_callback: str | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit setter and/or output compute functions for a binding manifest."""
    series_list = [s for s in bindings.get("series", []) if isinstance(s, dict)]
    emit_input = any(has_input_direction(s) for s in series_list)
    emit_output = any(has_output_direction(s) for s in series_list)
    if not emit_input and not emit_output:
        return []

    lines: list[str] = []
    include_aliases = True
    if emit_input:
        lines.extend(
            emit_setters_block(
                graph,
                workbook,
                bindings,
                export_addresses=export_addresses,
                include_type_aliases=include_aliases,
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
