"""Unit tests for view-level series binding groups (schema, ordering, manifest)."""

from __future__ import annotations

from typing import Any

import pytest

from excel_grapher.series_bindings.groups import (
    bindings_export_order,
    bindings_have_groups,
    group_manifest,
)
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    validate_bindings_document,
)


def _series(series_id: str, **overrides: Any) -> dict[str, Any]:
    """Minimal scalar binding entry with an input setter."""
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "Inputs",
        "data_range": "Inputs!B2",
        "layout": "scalar",
        "setter": {"name": f"set_{series_id}"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [],
        },
        "key": [],
    }
    entry.update(overrides)
    return entry


def _doc(*series: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": "1.5.0", "series": list(series)}


# --- Schema validation ---


def test_schema_accepts_groups_with_nested_path_and_order() -> None:
    doc = _doc(
        _series(
            "paris_debt",
            groups=[{"path": ["Climate scenarios", "Paris"], "order": 1}],
        )
    )
    validated = validate_bindings_document(doc)
    assert validated["series"][0]["groups"] == [
        {"path": ["Climate scenarios", "Paris"], "order": 1}
    ]


def test_schema_accepts_group_without_order() -> None:
    doc = _doc(_series("baseline_gdp", groups=[{"path": ["Baseline setup"]}]))
    validated = validate_bindings_document(doc)
    assert validated["series"][0]["groups"][0]["path"] == ["Baseline setup"]


def test_schema_rejects_empty_group_path() -> None:
    doc = _doc(_series("bad", groups=[{"path": []}]))
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_group_without_path() -> None:
    doc = _doc(_series("bad", groups=[{"order": 1}]))
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_non_string_path_segments() -> None:
    doc = _doc(_series("bad", groups=[{"path": [1, 2]}]))
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_non_integer_order() -> None:
    doc = _doc(_series("bad", groups=[{"path": ["G"], "order": "first"}]))
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_still_accepts_bindings_without_groups() -> None:
    doc = _doc(_series("plain"))
    validated = validate_bindings_document(doc)
    assert "groups" not in validated["series"][0]


# --- Ordering ---


def test_have_groups_detects_any_group_declaration() -> None:
    assert not bindings_have_groups(_doc(_series("a"), _series("b")))
    assert bindings_have_groups(_doc(_series("a"), _series("b", groups=[{"path": ["G"]}])))


def test_export_order_without_groups_preserves_declaration_order() -> None:
    doc = _doc(_series("zulu"), _series("alpha"))
    ordered = bindings_export_order(doc)
    assert [s["id"] for s in ordered] == ["zulu", "alpha"]


def test_export_order_sequences_by_group_first_appearance() -> None:
    doc = _doc(
        _series("c1", groups=[{"path": ["Climate"]}]),
        _series("b1", groups=[{"path": ["Baseline"]}]),
        _series("c2", groups=[{"path": ["Climate"]}]),
        _series("b2", groups=[{"path": ["Baseline"]}]),
    )
    ordered = bindings_export_order(doc)
    assert [s["id"] for s in ordered] == ["c1", "c2", "b1", "b2"]


def test_export_order_respects_order_within_leaf_group() -> None:
    doc = _doc(
        _series("second", groups=[{"path": ["G"], "order": 2}]),
        _series("first", groups=[{"path": ["G"], "order": 1}]),
        _series("unordered", groups=[{"path": ["G"]}]),
    )
    ordered = bindings_export_order(doc)
    # Explicit orders sort first; members without order fall back to declaration order.
    assert [s["id"] for s in ordered] == ["first", "second", "unordered"]


def test_export_order_emits_parent_members_before_nested_children() -> None:
    doc = _doc(
        _series("nested", groups=[{"path": ["Climate", "Paris"]}]),
        _series("parent_level", groups=[{"path": ["Climate"]}]),
    )
    ordered = bindings_export_order(doc)
    assert [s["id"] for s in ordered] == ["parent_level", "nested"]


def test_export_order_places_ungrouped_bindings_last() -> None:
    doc = _doc(
        _series("loose_a"),
        _series("grouped", groups=[{"path": ["G"]}]),
        _series("loose_b"),
    )
    ordered = bindings_export_order(doc)
    assert [s["id"] for s in ordered] == ["grouped", "loose_a", "loose_b"]


def test_export_order_uses_first_group_ref_for_placement() -> None:
    doc = _doc(
        _series("multi", groups=[{"path": ["A"]}, {"path": ["B"]}]),
        _series("b_only", groups=[{"path": ["B"]}]),
    )
    ordered = bindings_export_order(doc)
    assert [s["id"] for s in ordered] == ["multi", "b_only"]


# --- Manifest ---


def test_group_manifest_nests_children_and_lists_members() -> None:
    doc = _doc(
        _series("paris", groups=[{"path": ["Climate scenarios", "Paris"], "order": 1}]),
        _series("hot", groups=[{"path": ["Climate scenarios", "Hot"], "order": 2}]),
        _series("baseline_gdp", groups=[{"path": ["Baseline setup"]}]),
        _series("loose"),
    )
    manifest = group_manifest(doc)

    assert [g["label"] for g in manifest["groups"]] == [
        "Climate scenarios",
        "Baseline setup",
    ]
    climate = manifest["groups"][0]
    assert climate["path"] == ["Climate scenarios"]
    assert climate["slug"] == "climate_scenarios"
    assert climate["members"] == []
    assert [child["label"] for child in climate["children"]] == ["Paris", "Hot"]
    paris = climate["children"][0]
    assert paris["path"] == ["Climate scenarios", "Paris"]
    assert paris["members"] == [{"id": "paris", "setter": "set_paris", "compute": None, "order": 1}]
    assert manifest["ungrouped"] == [
        {"id": "loose", "setter": "set_loose", "compute": None, "order": None}
    ]


def test_group_manifest_records_multi_membership_in_every_group() -> None:
    doc = _doc(
        _series("multi", groups=[{"path": ["A"]}, {"path": ["B"], "order": 5}]),
    )
    manifest = group_manifest(doc)
    group_a, group_b = manifest["groups"]
    assert [m["id"] for m in group_a["members"]] == ["multi"]
    assert group_b["members"] == [
        {"id": "multi", "setter": "set_multi", "compute": None, "order": 5}
    ]


def test_group_manifest_includes_compute_names() -> None:
    entry = _series("outputs_debt", groups=[{"path": ["Outputs"]}])
    del entry["setter"]
    entry["output"] = {"compute": {"name": "compute_outputs_debt"}}
    manifest = group_manifest(_doc(entry))
    assert manifest["groups"][0]["members"] == [
        {"id": "outputs_debt", "setter": None, "compute": "compute_outputs_debt", "order": None}
    ]
