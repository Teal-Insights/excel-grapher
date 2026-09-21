"""Shared workbook and sidecar helpers for binding-authoring tests."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

import fastpyxl
import yaml

from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION

OBS_MEASURE: dict[str, Any] = {
    "concept": "OBS_VALUE",
    "dtype": "float",
    "bind": {"kind": "data_cell", "read": "float"},
}


def write_authoring_workbook(
    path: Path,
    *,
    sparse_year: bool = False,
    extra_engine_row: bool = False,
) -> Path:
    """Write a small Inputs / Engine / Outputs workbook for authoring tests."""
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = fastpyxl.Workbook()
    workbook.remove(workbook.active)
    inputs = workbook.create_sheet("Inputs")
    engine = workbook.create_sheet("Engine")
    outputs = workbook.create_sheet("Outputs")
    inputs["A1"] = 10
    inputs["B1"] = 0
    engine["B1"] = 2020
    engine["C1"] = None if sparse_year else 2021
    engine["A2"] = "Base"
    engine["B2"] = "=Inputs!A1+Inputs!B1+1"
    engine["C2"] = "=Inputs!A1+Inputs!B1+1"
    if extra_engine_row:
        engine["A3"] = "Alt"
        engine["B3"] = "=Engine!B2*2"
        engine["C3"] = "=Engine!C2*2"
    outputs["B1"] = "=Engine!B2"
    outputs["C1"] = "=Engine!C2"
    workbook.save(path)
    workbook.close()
    return path


def scalar_series(
    series_id: str,
    address: str,
    *,
    direction: str,
    compute_name: str | None = None,
) -> dict[str, Any]:
    """Return a keyless scalar series document for `address` (`Sheet!A1`)."""
    sheet, _sep, cell = address.partition("!")
    series: dict[str, Any] = {
        "id": series_id,
        "sheet": sheet,
        "data_range": address,
        "layout": "scalar",
        "structure": {"measure": dict(OBS_MEASURE), "dimensions": []},
        "key": [],
    }
    if direction == "output":
        name = compute_name or f"compute_{series_id}"
        series["output"] = {"compute": {"name": name, "record_contract": "records"}}
    else:
        series[direction] = {}
    return series


def years_internal_series(
    *,
    series_id: str = "engine_years",
    data_range: str = "Engine!B2:C2",
    fill: bool = False,
) -> dict[str, Any]:
    """Return a TIME_PERIOD `column_header` internal series over engine years."""
    bind: dict[str, Any] = {"kind": "column_header", "header_row": 1, "read": "int"}
    if fill:
        bind["fill"] = True
    return {
        "id": series_id,
        "sheet": "Engine",
        "data_range": data_range,
        "layout": "series",
        "internal": {},
        "structure": {
            "measure": dict(OBS_MEASURE),
            "dimensions": [
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": bind,
                }
            ],
        },
        "key": ["TIME_PERIOD"],
        "validation": {"warn_on_partial_overlap": False},
    }


def public_io_series() -> tuple[dict[str, Any], ...]:
    """Return the default public input/output series for the authoring workbook."""
    return (
        scalar_series("input_rate", "Inputs!A1", direction="input"),
        scalar_series("result_a", "Outputs!B1", direction="output"),
        scalar_series("result_b", "Outputs!C1", direction="output"),
    )


def binding_document(
    series: Sequence[Mapping[str, Any]],
    *,
    schema_version: str | None = None,
    workbook: str | None = "workbook.xlsx",
) -> dict[str, Any]:
    """Return a sidecar document wrapping `series`."""
    document: dict[str, Any] = {
        "schema_version": schema_version or CURRENT_SCHEMA_VERSION,
        "series": [dict(entry) for entry in series],
    }
    if workbook is not None:
        document["workbook"] = workbook
    return document


def write_shards(
    directory: Path,
    *,
    inputs: Sequence[Mapping[str, Any]] = (),
    outputs: Sequence[Mapping[str, Any]] = (),
    internals: Sequence[Mapping[str, Any]] = (),
    constants: Sequence[Mapping[str, Any]] = (),
    schema_version: str | None = None,
    workbook: str | None = "workbook.xlsx",
) -> Path:
    """Write the four conventional sidecar shards under `directory`."""
    directory.mkdir(parents=True, exist_ok=True)
    version = schema_version or CURRENT_SCHEMA_VERSION
    for name, entries in (
        ("inputs.bindings.yaml", inputs),
        ("outputs.bindings.yaml", outputs),
        ("internals.bindings.yaml", internals),
        ("constants.bindings.yaml", constants),
    ):
        (directory / name).write_text(
            yaml.safe_dump(
                binding_document(entries, schema_version=version, workbook=workbook),
                sort_keys=False,
                allow_unicode=True,
            ),
            encoding="utf-8",
        )
    return directory
