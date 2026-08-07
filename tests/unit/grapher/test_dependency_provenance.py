"""DependencyCause IntFlag bitmask and EdgeProvenance merge semantics."""

from __future__ import annotations

from enum import IntFlag

from excel_grapher.grapher.cache import _edge_provenance_from_json, _edge_provenance_to_json
from excel_grapher.grapher.dependency_provenance import (
    DependencyCause,
    EdgeProvenance,
    merge_edge_provenance,
)


def test_dependency_cause_is_int_flag() -> None:
    assert issubclass(DependencyCause, IntFlag)


def test_causes_combine_with_bitwise_or() -> None:
    combined = DependencyCause.direct_ref | DependencyCause.static_range
    assert DependencyCause.direct_ref in combined
    assert DependencyCause.static_range in combined
    assert DependencyCause.dynamic_offset not in combined


def test_edge_provenance_stores_causes_as_flag_not_frozenset() -> None:
    prov = EdgeProvenance(causes=DependencyCause.direct_ref)
    assert prov.causes is DependencyCause.direct_ref
    assert isinstance(prov.causes, DependencyCause)


def test_empty_provenance_has_no_causes() -> None:
    empty = EdgeProvenance.empty()
    assert empty.causes == DependencyCause(0)
    assert DependencyCause.direct_ref not in empty.causes


def test_merge_unions_cause_flags() -> None:
    a = EdgeProvenance(causes=DependencyCause.direct_ref)
    b = EdgeProvenance(causes=DependencyCause.dynamic_offset)
    merged = a.merge(b)
    assert merged.causes == DependencyCause.direct_ref | DependencyCause.dynamic_offset


def test_merge_edge_provenance_none_passthrough() -> None:
    a = EdgeProvenance(causes=DependencyCause.static_range)
    assert merge_edge_provenance(None, a) is a
    assert merge_edge_provenance(a, None) is a
    assert merge_edge_provenance(None, None) is None


def test_json_round_trip_preserves_cause_names() -> None:
    prov = EdgeProvenance(
        causes=DependencyCause.direct_ref | DependencyCause.static_range,
        direct_sites_normalized=((1, 11),),
    )
    blob = _edge_provenance_to_json(prov)
    assert blob["causes"] == ["direct_ref", "static_range"]
    restored = _edge_provenance_from_json(blob)
    assert restored.causes == prov.causes
    assert restored.direct_sites_normalized == ((1, 11),)


def test_json_from_legacy_string_cause_list() -> None:
    restored = _edge_provenance_from_json(
        {
            "causes": ["direct_ref", "dynamic_indirect"],
            "direct_sites_formula": [],
            "direct_sites_normalized": [],
        }
    )
    assert DependencyCause.direct_ref in restored.causes
    assert DependencyCause.dynamic_indirect in restored.causes
    assert DependencyCause.static_range not in restored.causes
