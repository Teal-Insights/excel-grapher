"""Derive constant series from explicit series binding manifests."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.derive_common import derive_series_for_direction
from excel_grapher.series_bindings.normalize import has_constant_direction
from excel_grapher.series_bindings.types import (
    ConstantSeries,
    ConstantSeriesCell,
    SeriesResolution,
    WorkbookSeriesBindings,
)


def _build_constant_series(
    resolved: SeriesResolution,
    series: dict[str, Any],
    cells: list[ConstantSeriesCell],
) -> ConstantSeries:
    return {
        "id": resolved["series_id"],
        "key_fields": [str(field) for field in (series.get("key") or [])],
        "requires_address": resolved["requires_address"],
        "cells": cells,
        "issues": resolved["issues"],
    }


def derive_constant_series(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> list[ConstantSeries]:
    """Return constant series for graph-leaf bindings.

    Each manifest `series[]` entry with `constant: {}` resolves to per-cell
    `{address, key, record}` for graph leaves in `data_range`. Package export
    bakes these as data constants rather than `compute_*` arguments.
    """
    return derive_series_for_direction(
        graph,
        bindings,
        workbook=workbook,
        direction="constant",
        has_direction=has_constant_direction,
        build_series=_build_constant_series,
    )
