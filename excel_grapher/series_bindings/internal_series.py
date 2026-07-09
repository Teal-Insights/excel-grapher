"""Derive internal series from explicit series binding manifests."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import has_internal_direction
from excel_grapher.series_bindings.resolve import resolve_series_bindings
from excel_grapher.series_bindings.types import (
    InternalSeries,
    InternalSeriesCell,
    WorkbookSeriesBindings,
)


def derive_internal_series(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> list[InternalSeries]:
    """Return one internal series per binding entry with internal direction and graph overlap."""
    report = resolve_series_bindings(graph, bindings, workbook=workbook, direction="internal")
    series_by_id = {
        str(series["id"]): series
        for series in bindings.get("series", [])
        if isinstance(series, dict) and "id" in series and has_internal_direction(series)
    }

    internal_series: list[InternalSeries] = []
    for resolved in report["series"]:
        if not resolved["leaves"]:
            continue
        series = series_by_id.get(resolved["series_id"])
        if series is None:
            continue
        cells: list[InternalSeriesCell] = [
            {
                "address": leaf["address"],
                "coordinates": leaf["coordinates"],
                "key": leaf["key"],
                "record": leaf["record"],
            }
            for leaf in resolved["leaves"]
        ]
        internal_series.append(
            {
                "id": resolved["series_id"],
                "key_fields": [str(field) for field in (series.get("key") or [])],
                "requires_address": resolved["requires_address"],
                "cells": cells,
                "issues": resolved["issues"],
            }
        )
    return internal_series
