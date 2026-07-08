"""Derive input series from explicit series binding manifests."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import has_input_direction
from excel_grapher.series_bindings.resolve import resolve_series_bindings
from excel_grapher.series_bindings.types import (
    InputSeries,
    InputSeriesCell,
    WorkbookSeriesBindings,
)


def _setter_name(series: dict[str, Any]) -> str:
    input_block = series.get("input") or {}
    setter = input_block.get("setter") or series.get("setter") or {}
    return str(setter.get("name", f"set_{series.get('id', 'series')}"))


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
    report = resolve_series_bindings(graph, bindings, workbook=workbook, direction="input")
    series_by_id = {
        str(series["id"]): series
        for series in bindings.get("series", [])
        if isinstance(series, dict) and "id" in series and has_input_direction(series)
    }

    input_series: list[InputSeries] = []
    for resolved in report["series"]:
        if not resolved["leaves"]:
            continue
        series = series_by_id.get(resolved["series_id"])
        if series is None:
            continue
        cells: list[InputSeriesCell] = [
            {
                "address": leaf["address"],
                "coordinates": leaf["coordinates"],
                "key": leaf["key"],
                "record": leaf["record"],
            }
            for leaf in resolved["leaves"]
        ]
        input_series.append(
            {
                "id": resolved["series_id"],
                "setter_name": _setter_name(series),
                "key_fields": [str(field) for field in (series.get("key") or [])],
                "requires_address": resolved["requires_address"],
                "cells": cells,
                "issues": resolved["issues"],
            }
        )
    return input_series
