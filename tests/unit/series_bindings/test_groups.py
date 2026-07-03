"""Unit tests for series binding API group ordering and manifest."""

from __future__ import annotations

from typing import Any

from excel_grapher.series_bindings.groups import (
    any_series_has_groups,
    build_api_group_manifest,
    ordered_compute_names,
    ordered_series_for_direction,
    ordered_setter_names,
)
from excel_grapher.series_bindings.types import WorkbookSeriesBindings


def _bindings(*series: dict[str, Any]) -> WorkbookSeriesBindings:
    return {"schema_version": "1.6.0", "series": list(series)}


def _setter_series(
    series_id: str,
    setter_name: str,
    *,
    groups: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "S",
        "data_range": "S!A1",
        "input": {"setter": {"name": setter_name}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
            "dimensions": [],
        },
        "key": [],
        "layout": "scalar",
    }
    if groups is not None:
        entry["groups"] = groups
    return entry


def _compute_series(
    series_id: str,
    compute_name: str,
    *,
    groups: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "S",
        "data_range": "S!A1",
        "output": {"compute": {"name": compute_name}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
            "dimensions": [],
        },
        "key": [],
        "layout": "scalar",
    }
    if groups is not None:
        entry["groups"] = groups
    return entry


def test_any_series_has_groups_false_when_all_omit_groups() -> None:
    bindings = _bindings(
        _setter_series("a", "set_a"),
        _setter_series("b", "set_b"),
    )
    assert any_series_has_groups(bindings) is False


def test_any_series_has_groups_true_when_one_declares_groups() -> None:
    bindings = _bindings(
        _setter_series("a", "set_a"),
        _setter_series("b", "set_b", groups=[{"path": ["Macro"]}]),
    )
    assert any_series_has_groups(bindings) is True


def test_ordered_setter_names_alpha_when_no_groups() -> None:
    bindings = _bindings(
        _setter_series("z", "set_z"),
        _setter_series("a", "set_a"),
    )
    assert ordered_setter_names(bindings) == ["set_a", "set_z"]


def test_ordered_setter_names_by_group_path_and_order() -> None:
    bindings = _bindings(
        _setter_series(
            "paris",
            "set_paris",
            groups=[{"path": ["Climate scenarios", "Paris"], "order": 2}],
        ),
        _setter_series(
            "baseline",
            "set_baseline",
            groups=[{"path": ["Baseline setup"], "order": 1}],
        ),
        _setter_series(
            "moderate",
            "set_moderate",
            groups=[{"path": ["Climate scenarios", "Moderate"], "order": 1}],
        ),
        _setter_series("ungrouped", "set_ungrouped"),
    )
    assert ordered_setter_names(bindings) == [
        "set_baseline",
        "set_moderate",
        "set_paris",
        "set_ungrouped",
    ]


def test_ordered_setter_names_uses_manifest_order_within_group_when_order_omitted() -> None:
    bindings = _bindings(
        _setter_series("second", "set_second", groups=[{"path": ["Macro"]}]),
        _setter_series("first", "set_first", groups=[{"path": ["Macro"]}]),
    )
    assert ordered_setter_names(bindings) == ["set_second", "set_first"]


def test_ordered_compute_names_respects_groups() -> None:
    bindings = _bindings(
        _compute_series("b", "compute_b", groups=[{"path": ["Outputs"], "order": 2}]),
        _compute_series("a", "compute_a", groups=[{"path": ["Outputs"], "order": 1}]),
    )
    assert ordered_compute_names(bindings) == ["compute_a", "compute_b"]


def test_ordered_series_for_direction_filters_direction() -> None:
    bindings = _bindings(
        _setter_series("in_only", "set_in_only", groups=[{"path": ["Inputs"]}]),
        _compute_series("out_only", "compute_out_only", groups=[{"path": ["Outputs"]}]),
    )
    input_ids = [s["id"] for s in ordered_series_for_direction(bindings, "input")]
    output_ids = [s["id"] for s in ordered_series_for_direction(bindings, "output")]
    assert input_ids == ["in_only"]
    assert output_ids == ["out_only"]


def test_build_api_group_manifest_nested_tree() -> None:
    bindings = _bindings(
        _setter_series(
            "paris",
            "set_paris",
            groups=[{"path": ["Climate scenarios", "Paris"], "order": 1}],
        ),
        _compute_series(
            "paris_out",
            "compute_paris",
            groups=[{"path": ["Climate scenarios", "Paris"], "order": 2}],
        ),
        _setter_series("baseline", "set_baseline", groups=[{"path": ["Baseline setup"]}]),
    )
    manifest = build_api_group_manifest(bindings)
    assert manifest["schema_version"] == "1.0.0"
    assert manifest["flat"]["setters"] == ordered_setter_names(bindings)
    assert manifest["flat"]["computes"] == ordered_compute_names(bindings)
    climate = manifest["group_tree"]["Climate scenarios"]
    assert climate["Paris"]["setters"] == ["set_paris"]
    assert climate["Paris"]["computes"] == ["compute_paris"]
    assert manifest["group_tree"]["Baseline setup"]["setters"] == ["set_baseline"]
