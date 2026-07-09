"""Derive internal series from explicit series binding manifests."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.derive_common import derive_series_for_direction
from excel_grapher.series_bindings.normalize import has_internal_direction
from excel_grapher.series_bindings.types import (
    InternalSeries,
    InternalSeriesCell,
    SeriesResolution,
    WorkbookSeriesBindings,
)


def _build_internal_series(
    resolved: SeriesResolution,
    series: dict[str, Any],
    cells: list[InternalSeriesCell],
) -> InternalSeries:
    return {
        "id": resolved["series_id"],
        "key_fields": [str(field) for field in (series.get("key") or [])],
        "requires_address": resolved["requires_address"],
        "cells": cells,
        "issues": resolved["issues"],
    }


def derive_internal_series(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> list[InternalSeries]:
    """Return internal series for formula-cell key triangulation without public I/O APIs.

    Each manifest `series[]` entry with `internal: {}` resolves to per-cell
    `{address, key, record}` for formula nodes in `data_range`. No `set_*` or
    `compute_*` codegen is emitted for these entries.
    """
    return derive_series_for_direction(
        graph,
        bindings,
        workbook=workbook,
        direction="internal",
        has_direction=has_internal_direction,
        build_series=_build_internal_series,
    )
