"""Sparse coordinate leaf store (#579)."""

from __future__ import annotations

from types import MappingProxyType
from typing import Any, cast

import pytest

from excel_grapher.runtime.leaves import (
    MISSING,
    LeafInputs,
    as_leaf_store,
    leaf,
    overlay_leaf_inputs,
    prepare_context_inputs,
)


def test_as_leaf_store_converts_nodekey_dict() -> None:
    store = as_leaf_store({"Sheet1!A1": 10.0, "Sheet1!B2": "hi"})
    assert store == {"Sheet1": {(1, 1): 10.0, (2, 2): "hi"}}


def test_as_leaf_store_copies_nested_store() -> None:
    original = {"Sheet1": {(1, 1): 1}}
    store = as_leaf_store(original)
    store["Sheet1"][(1, 1)] = 99
    assert original["Sheet1"][(1, 1)] == 1


def test_as_leaf_store_interns_sheet_names() -> None:
    store = as_leaf_store({"Sheet1!A1": 1, "Sheet1!A2": 2, "Other!A1": 3})
    assert set(store) == {"Sheet1", "Other"}
    assert store["Sheet1"][(1, 1)] == 1
    assert store["Sheet1"][(2, 1)] == 2
    assert store["Other"][(1, 1)] == 3


def test_leaf_returns_missing_for_absent_coords() -> None:
    store = {"Sheet1": {(1, 1): 0}}
    assert leaf(store, "Sheet1", 1, 1) == 0
    assert leaf(store, "Sheet1", 1, 2) is MISSING
    assert leaf(store, "Missing", 1, 1) is MISSING


def test_overlay_nodekeys_wins_over_constants() -> None:
    store = as_leaf_store({"Sheet1!A1": 1, "Sheet1!A2": 2})
    overlay_leaf_inputs(store, {"Sheet1!A1": 99})
    assert store["Sheet1"][(1, 1)] == 99
    assert store["Sheet1"][(2, 1)] == 2


def test_overlay_fails_closed_on_unparseable_key() -> None:
    store: dict[str, dict[tuple[int, int], object]] = {}
    with pytest.raises(ValueError, match="Cannot round-trip"):
        overlay_leaf_inputs(store, {"A1": 1})


def test_prepare_context_inputs_merges_constants_then_overlay() -> None:
    defaults = {"Sheet1": {(1, 1): 1}}
    constants = {"Sheet1": {(2, 1): 2}}
    merged = prepare_context_inputs(defaults, constants, {"Sheet1!A1": 9})
    assert merged == {"Sheet1": {(1, 1): 9, (2, 1): 2}}
    assert defaults["Sheet1"][(1, 1)] == 1


def test_prepare_context_inputs_accepts_mapping_proxy_constants() -> None:
    defaults = {"Sheet1": {(1, 1): 1}}
    constants = MappingProxyType({"Sheet1": MappingProxyType({(2, 1): 2})})
    merged = prepare_context_inputs(defaults, constants, {"Sheet1!A1": 9})
    assert merged == {"Sheet1": {(1, 1): 9, (2, 1): 2}}
    assert defaults["Sheet1"][(1, 1)] == 1
    with pytest.raises(TypeError):
        cast(Any, constants)["Sheet1"] = {}


def test_leaf_inputs_nodekey_view() -> None:
    store = as_leaf_store({"Sheet1!A1": 5})
    view = LeafInputs(store)
    assert "Sheet1!A1" in view
    assert view["Sheet1!A1"] == 5
    view["Sheet1!B1"] = 7
    assert store["Sheet1"][(1, 2)] == 7
    assert view.get("Sheet1!Z9") is None
