"""Derive input series from explicit series binding manifests."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.derive_common import derive_series_for_direction
from excel_grapher.series_bindings.normalize import has_input_direction
from excel_grapher.series_bindings.types import (
    InputSeries,
    InputSeriesCell,
    SeriesResolution,
    WorkbookSeriesBindings,
)


def _setter_name(series: dict[str, Any]) -> str:
    input_block = series.get("input") or {}
    setter = input_block.get("setter") or series.get("setter") or {}
    return str(setter.get("name", f"set_{series.get('id', 'series')}"))


def _build_input_series(
    resolved: SeriesResolution,
    series: dict[str, Any],
    cells: list[InputSeriesCell],
) -> InputSeries:
    return {
        "id": resolved["series_id"],
        "setter_name": _setter_name(series),
        "key_fields": [str(field) for field in (series.get("key") or [])],
        "requires_address": resolved["requires_address"],
        "cells": cells,
        "issues": resolved["issues"],
    }


def derive_input_series(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> list[InputSeries]:
    """Return one input series per binding series with graph-leaf or override overlap.

    Series bindings are the semantic source of truth: each input series
    corresponds to one manifest `series[]` entry, and each cell corresponds to
    a resolved graph leaf or override cell participating in that series.
    """
    return derive_series_for_direction(
        graph,
        bindings,
        workbook=workbook,
        direction="input",
        has_direction=has_input_direction,
        build_series=_build_input_series,
    )
