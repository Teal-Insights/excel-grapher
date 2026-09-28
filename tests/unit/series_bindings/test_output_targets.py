"""Tests for deriving graph extraction targets from output series bindings."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
from fastpyxl.workbook.defined_name import DefinedName

from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from excel_grapher.series_bindings.workflow import (
    output_series_targets,
    target_set_sha256,
)


def _workbook(tmp_path: Path) -> Path:
    path = tmp_path / "model.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Model"
    for col in "ABC":
        ws[f"{col}1"] = 1
        for row in (2, 3, 4):
            ws[f"{col}{row}"] = f"={col}1*{row}"
    wb.defined_names.add(DefinedName(name="Doubled", attr_text="Model!$A$2:$C$2"))
    wb.save(path)
    return path


def _bindings(*series: dict) -> WorkbookSeriesBindings:
    return {"schema_version": CURRENT_SCHEMA_VERSION, "series": list(series)}


_INPUT = {"id": "growth", "sheet": "Model", "data_range": "Model!A1:C1", "input": {}}


def test_output_series_targets_skips_non_output_series_and_applies_exclude_columns(
    tmp_path: Path,
) -> None:
    workbook = _workbook(tmp_path)
    bindings = _bindings(
        _INPUT,
        {
            "id": "doubled",
            "sheet": "Model",
            "data_range": "Model!A2:C2",
            "exclude_columns": ["B"],
            "output": {},
        },
        {"id": "mid", "sheet": "Model", "data_range": "Model!A3:C3", "internal": {}},
        {"id": "k", "sheet": "Model", "data_range": "Model!A4", "constant": {}},
    )

    assert output_series_targets(bindings, workbook=workbook) == ["Model!A2", "Model!C2"]


def test_output_series_targets_applies_exclude_rows(tmp_path: Path) -> None:
    workbook = _workbook(tmp_path)
    bindings = _bindings(
        {
            "id": "block",
            "sheet": "Model",
            "data_range": "Model!A2:A4",
            "exclude_rows": [3],
            "output": {},
        }
    )

    assert output_series_targets(bindings, workbook=workbook) == ["Model!A2", "Model!A4"]


def test_output_series_targets_expands_named_range(tmp_path: Path) -> None:
    workbook = _workbook(tmp_path)
    bindings = _bindings({"id": "doubled", "sheet": "Model", "data_range": "Doubled", "output": {}})

    assert output_series_targets(bindings, workbook=workbook) == [
        "Model!A2",
        "Model!B2",
        "Model!C2",
    ]


def test_output_series_targets_expands_multi_range_series_sorted_and_unique(
    tmp_path: Path,
) -> None:
    workbook = _workbook(tmp_path)
    bindings = _bindings(
        {
            "id": "split",
            "sheet": "Model",
            "data_range": ["Model!C4", "Model!A2:A3", "Model!A3"],
            "output": {},
        },
        {"id": "again", "sheet": "Model", "data_range": "Model!A2", "output": {}},
    )

    assert output_series_targets(bindings, workbook=workbook) == [
        "Model!A2",
        "Model!A3",
        "Model!C4",
    ]


def test_output_series_targets_unions_extra_targets(tmp_path: Path) -> None:
    workbook = _workbook(tmp_path)
    bindings = _bindings({"id": "a", "sheet": "Model", "data_range": "Model!B2", "output": {}})

    assert output_series_targets(
        bindings, workbook=workbook, extra_targets=["Model!C4", "Model!B2"]
    ) == ["Model!B2", "Model!C4"]


def test_target_set_sha256_ignores_order_duplicates_and_non_output_edits(
    tmp_path: Path,
) -> None:
    workbook = _workbook(tmp_path)
    output = {"id": "doubled", "sheet": "Model", "data_range": "Model!A2:C2", "output": {}}
    before = output_series_targets(_bindings(_INPUT, output), workbook=workbook)
    edited_input = {**_INPUT, "data_range": "Model!A1:B1"}
    after = output_series_targets(_bindings(edited_input, output), workbook=workbook)

    assert target_set_sha256(before) == target_set_sha256(after)
    assert target_set_sha256(["Model!B2", "Model!A2", "Model!A2"]) == target_set_sha256(
        ["Model!A2", "Model!B2"]
    )
    assert target_set_sha256(["Model!A2"]) != target_set_sha256(["Model!A3"])
