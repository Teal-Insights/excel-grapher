"""Derive output series from explicit series binding manifests."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.derive_common import derive_series_for_direction
from excel_grapher.series_bindings.normalize import has_output_direction
from excel_grapher.series_bindings.types import (
    OutputSeries,
    OutputSeriesCell,
    SeriesResolution,
    WorkbookSeriesBindings,
)


def _compute_name(series: dict[str, Any]) -> str:
    output = series.get("output") or {}
    compute = output.get("compute") or {}
    return str(compute.get("name", f"compute_{series.get('id', 'series')}"))


def _build_output_series(
    resolved: SeriesResolution,
    series: dict[str, Any],
    cells: list[OutputSeriesCell],
) -> OutputSeries:
    return {
        "id": resolved["series_id"],
        "compute_name": _compute_name(series),
        "key_fields": [str(field) for field in (series.get("key") or [])],
        "cells": cells,
        "issues": resolved["issues"],
    }


def derive_output_series(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> list[OutputSeries]:
    """Return one output series per binding entry with output.compute and graph overlap."""
    return derive_series_for_direction(
        graph,
        bindings,
        workbook=workbook,
        direction="output",
        has_direction=has_output_direction,
        build_series=_build_output_series,
    )
