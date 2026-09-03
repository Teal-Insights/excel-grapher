"""Symbolic IndexSet algebra: range, slice, affine, predicate, irregular gather."""

from __future__ import annotations

import pytest

from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import DependenceEdge, IndexSet


def test_from_indices_compresses_contiguous_range() -> None:
    assert IndexSet.from_indices((0, 1, 2, 3)) == IndexSet.interval(0, 4)
    assert IndexSet.from_indices(range(10)) == IndexSet.interval(0, 10)
    assert IndexSet.from_indices(()) == IndexSet.empty()
    assert IndexSet.from_indices((7,)) == IndexSet.interval(7, 8)


def test_from_indices_compresses_strided_slice() -> None:
    assert IndexSet.from_indices((0, 2, 4, 6)) == IndexSet.interval(0, 8, 2)
    assert IndexSet.from_indices((1, 4, 7)) == IndexSet.interval(1, 8, 3)


def test_from_indices_keeps_irregular_gather() -> None:
    punched = IndexSet.from_indices((0, 2, 5))
    assert punched.materialize() == (0, 2, 5)
    assert punched.to_source() == "(0, 2, 5)"
    assert not punched.is_progression()


def test_union_merges_adjacent_and_overlapping_ranges() -> None:
    assert IndexSet.interval(0, 3).union(IndexSet.interval(3, 5)) == IndexSet.interval(0, 5)
    assert IndexSet.interval(0, 3).union(IndexSet.interval(1, 4)) == IndexSet.interval(0, 4)
    assert IndexSet.interval(0, 8, 2).union(IndexSet.interval(1, 9, 2)) == IndexSet.interval(0, 8)


def test_union_of_gapped_ranges_stays_gathered() -> None:
    got = IndexSet.interval(0, 2).union(IndexSet.interval(4, 6))
    assert got.materialize() == (0, 1, 4, 5)
    assert got.to_source() == "(0, 1, 4, 5)"


def test_affine_image_of_range_is_strided() -> None:
    assert IndexSet.interval(0, 5).map_affine(2, 1) == IndexSet.interval(1, 11, 2)
    assert IndexSet.interval(0, 4).map_affine(1, -1).materialize() == (-1, 0, 1, 2)
    assert IndexSet.from_indices((0, 2)).map_affine(3, 1).materialize() == (1, 7)
    assert IndexSet.interval(0, 5).map_affine(-1, 10) == IndexSet.interval(6, 11)


def test_affine_constructor_normalizes() -> None:
    base = IndexSet.interval(0, 5)
    assert IndexSet.affine(base, 2, 1) == IndexSet.interval(1, 11, 2)
    assert IndexSet.affine(IndexSet.empty(), 2, 1) == IndexSet.empty()
    assert IndexSet.affine(IndexSet.interval(0, 3), 0, 4) == IndexSet.interval(4, 5)


def test_residue_predicate_compresses_to_slice() -> None:
    evens = IndexSet.interval(0, 10).filter_residue(2, 0)
    assert evens == IndexSet.interval(0, 10, 2)
    odds = IndexSet.interval(0, 10).filter_residue(2, 1)
    assert odds == IndexSet.interval(1, 10, 2)


def test_callable_predicate_materializes_then_compresses() -> None:
    got = IndexSet.interval(0, 8).filter(lambda i: i % 3 == 0)
    assert got == IndexSet.interval(0, 8, 3)
    punched = IndexSet.interval(0, 6).filter(lambda i: i in {0, 3, 4})
    assert punched.materialize() == (0, 3, 4)
    assert punched.to_source() == "(0, 3, 4)"


def test_closure_under_unit_lag_fills_prefix() -> None:
    assert IndexSet.from_indices((2, 4)).closure_under((1,)) == IndexSet.interval(0, 5)
    assert IndexSet.interval(2, 5).closure_under((1,)) == IndexSet.interval(0, 5)
    assert IndexSet.empty().closure_under((1,)) == IndexSet.empty()
    assert IndexSet.interval(0, 3).closure_under((1,)) == IndexSet.interval(0, 3)


def test_closure_under_stride_walks_that_distance() -> None:
    assert IndexSet.from_indices((4,)).closure_under((2,)) == IndexSet.interval(0, 5, 2)
    assert IndexSet.interval(4, 10, 2).closure_under((2,)) == IndexSet.interval(0, 10, 2)
    assert IndexSet.interval(5, 10, 2).closure_under((2,)) == IndexSet.interval(1, 10, 2)


def test_closure_under_ignores_non_positive_distances() -> None:
    s = IndexSet.from_indices((3,))
    assert s.closure_under((0, -1)) == s
    assert s.closure_under(()) == s


def test_closure_under_mixed_strides_may_gather() -> None:
    got = IndexSet.from_indices((7,)).closure_under((2, 3))
    assert got.materialize() == (0, 1, 2, 3, 4, 5, 7)


def test_closure_under_edges_uses_producer_lags() -> None:
    edges = (
        DependenceEdge("debt", "debt", "Engine!B2", "Engine!A2", 1),
        DependenceEdge("debt", "adjustment", "Engine!B2", "Engine!B3", 0),
        DependenceEdge("adjustment", "debt", "Engine!B3", "Engine!A2", 1),
    )
    demanded = IndexSet.from_indices((2, 4))
    assert demanded.closure_under_edges(edges, producer_id="debt") == IndexSet.interval(0, 5)
    assert demanded.closure_under_edges(edges, producer_id="adjustment") == demanded


def test_to_source_emits_range_slice_or_tuple() -> None:
    assert IndexSet.interval(0, 3).to_source() == "range(0, 3)"
    assert IndexSet.interval(1, 3).to_source() == "range(1, 3)"
    assert IndexSet.interval(0, 8, 2).to_source() == "range(0, 8, 2)"
    assert IndexSet.interval(7, 8).to_source() == "(7,)"
    assert IndexSet.empty().to_source() == "()"
    assert IndexSet.from_indices((0, 2, 5)).to_source() == "(0, 2, 5)"


def test_positions_in_contiguous_universe() -> None:
    wanted = IndexSet.interval(1, 3)
    computed = IndexSet.interval(0, 5)
    assert wanted.positions_in(computed) == IndexSet.interval(1, 3)
    shifted = IndexSet.interval(4, 6).positions_in(IndexSet.interval(2, 8))
    assert shifted == IndexSet.interval(2, 4)


def test_positions_in_strided_universe() -> None:
    wanted = IndexSet.interval(4, 8, 2)
    universe = IndexSet.interval(2, 10, 2)
    assert wanted.positions_in(universe) == IndexSet.interval(1, 3)


def test_positions_in_missing_index_is_export_error() -> None:
    with pytest.raises(InvertedTreeExportError, match="not in the universe"):
        IndexSet.interval(0, 2).positions_in(IndexSet.interval(3, 5))


def test_long_range_source_is_independent_of_member_count() -> None:
    s = IndexSet.interval(0, 200)
    assert s.to_source() == "range(0, 200)"
    assert len(s.to_source()) < 20
