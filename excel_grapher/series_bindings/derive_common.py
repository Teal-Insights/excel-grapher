"""Shared helpers for deriving typed series from binding resolution reports."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, TypeVar

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.resolve import BindingDirection, resolve_series_bindings
from excel_grapher.series_bindings.types import (
    InputSeriesCell,
    SeriesResolution,
    WorkbookSeriesBindings,
)

T = TypeVar("T")


def series_entries_by_id(
    bindings: WorkbookSeriesBindings,
    *,
    has_direction: Callable[[dict[str, Any]], bool],
) -> dict[str, dict[str, Any]]:
    """Index manifest series entries that declare the requested direction."""
    return {
        str(series["id"]): series
        for series in bindings.get("series", [])
        if isinstance(series, dict) and "id" in series and has_direction(series)
    }


def resolved_series_cells(resolved: SeriesResolution) -> list[InputSeriesCell]:
    """Map resolved leaves to the shared per-cell payload used by derive APIs."""
    return [
        {
            "address": leaf["address"],
            "coordinates": leaf["coordinates"],
            "key": leaf["key"],
            "record": leaf["record"],
        }
        for leaf in resolved["leaves"]
    ]


def derive_series_for_direction(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    direction: BindingDirection,
    has_direction: Callable[[dict[str, Any]], bool],
    build_series: Callable[[SeriesResolution, dict[str, Any], list[InputSeriesCell]], T],
) -> list[T]:
    """Resolve and materialize one typed series object per manifest entry."""
    report = resolve_series_bindings(graph, bindings, workbook=workbook, direction=direction)
    series_by_id = series_entries_by_id(bindings, has_direction=has_direction)

    derived: list[T] = []
    for resolved in report["series"]:
        if not resolved["leaves"]:
            continue
        series = series_by_id.get(resolved["series_id"])
        if series is None:
            continue
        derived.append(build_series(resolved, series, resolved_series_cells(resolved)))
    return derived
