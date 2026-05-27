"""Derive output series from explicit series binding manifests."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import has_output_direction
from excel_grapher.series_bindings.resolve import resolve_series_bindings
from excel_grapher.series_bindings.types import (
    OutputSeries,
    OutputSeriesCell,
    WorkbookSeriesBindings,
)


def _compute_name(series: dict[str, Any]) -> str:
    output = series.get("output") or {}
    compute = output.get("compute") or {}
    return str(compute.get("name", f"compute_{series.get('id', 'series')}"))


def derive_output_series(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> list[OutputSeries]:
    """Return one output series per binding entry with output.compute and graph overlap."""
    report = resolve_series_bindings(graph, bindings, workbook=workbook, direction="output")
    series_by_id = {
        str(series["id"]): series
        for series in bindings.get("series", [])
        if isinstance(series, dict) and "id" in series and has_output_direction(series)
    }

    output_series: list[OutputSeries] = []
    for resolved in report["series"]:
        if not resolved["leaves"]:
            continue
        series = series_by_id.get(resolved["series_id"])
        if series is None:
            continue
        cells: list[OutputSeriesCell] = [
            {
                "address": leaf["address"],
                "coordinates": leaf["coordinates"],
                "key": leaf["key"],
                "record": leaf["record"],
            }
            for leaf in resolved["leaves"]
        ]
        output_series.append(
            {
                "id": resolved["series_id"],
                "compute_name": _compute_name(series),
                "key_fields": [str(field) for field in (series.get("key") or [])],
                "cells": cells,
                "issues": resolved["issues"],
            }
        )
    return output_series
