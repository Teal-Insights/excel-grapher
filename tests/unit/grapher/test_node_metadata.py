"""Node metadata storage: shared empty singleton and accessor writes (#493)."""

from __future__ import annotations

import copy
import pickle
from collections.abc import Mapping, MutableMapping
from typing import Any, cast

import pytest

from excel_grapher.grapher.cache import dependency_graph_from_json, dependency_graph_to_json
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    EMPTY_METADATA,
    Node,
    copy_metadata,
    make_cell_node,
    node_to_view,
)
from excel_grapher.grapher.subgraph import select_path_induced_subgraph


def _leaf(sheet: str = "Sheet1", column: str = "A", row: int = 1, value: object = 1) -> Node:
    return make_cell_node(sheet, column, row, value=value, is_leaf=True)


def _formula_node(sheet: str = "Sheet1", column: str = "B", row: int = 1) -> Node:
    return make_cell_node(
        sheet,
        column,
        row,
        formula="=A1",
        normalized_formula=f"={sheet}!A1",
        is_leaf=False,
    )


# -------------------------------------------------------------------
# Empty metadata is a shared singleton, never a per-node dict
# -------------------------------------------------------------------


def test_fresh_nodes_share_the_empty_metadata_singleton() -> None:
    """No node allocates a dict for metadata it never uses."""
    cell = _leaf()
    other = make_cell_node("Sheet1", "E", 5)
    direct = Node(
        sheet="Sheet1",
        column="A",
        row=2,
        formula=None,
        normalized_formula=None,
        value=None,
        is_leaf=True,
    )
    explicit_empty = make_cell_node("Sheet1", "A", 3, metadata={})

    assert cell.metadata is EMPTY_METADATA
    assert other.metadata is EMPTY_METADATA
    assert direct.metadata is EMPTY_METADATA
    assert explicit_empty.metadata is EMPTY_METADATA


def test_empty_metadata_reads_as_an_empty_mapping() -> None:
    node = _leaf()
    assert isinstance(node.metadata, Mapping)
    assert len(node.metadata) == 0
    assert dict(node.metadata) == {}
    assert node.metadata == {}
    assert "k" not in node.metadata
    assert node.metadata.get("k") is None
    assert list(node.metadata.items()) == []
    assert repr(node.metadata) == "{}"
    with pytest.raises(KeyError):
        _ = node.metadata["k"]


def test_empty_metadata_rejects_item_assignment() -> None:
    """Writers go through `set_metadata` / `update_metadata`, not item assignment."""
    node = _leaf()
    with pytest.raises(TypeError):
        cast(MutableMapping[str, Any], node.metadata)["k"] = 1


def test_copy_metadata_shares_singleton_for_empty_input() -> None:
    assert copy_metadata(None) is EMPTY_METADATA
    assert copy_metadata({}) is EMPTY_METADATA
    assert copy_metadata(EMPTY_METADATA) is EMPTY_METADATA

    source: dict[str, Any] = {"k": 1}
    copied = copy_metadata(source)
    assert copied == {"k": 1}
    assert copied is not source


# -------------------------------------------------------------------
# Accessor writes
# -------------------------------------------------------------------


def test_set_metadata_copies_input_and_normalizes_empty() -> None:
    node = _leaf()
    source: dict[str, Any] = {"k": 1}
    node.set_metadata(source)
    source["k"] = 2
    assert dict(node.metadata) == {"k": 1}

    node.set_metadata({})
    assert node.metadata is EMPTY_METADATA

    node.set_metadata({"k": 1})
    node.set_metadata(None)
    assert node.metadata is EMPTY_METADATA


def test_update_metadata_upgrades_from_singleton_and_merges() -> None:
    node = _leaf()
    node.update_metadata({"a": 1})
    assert dict(node.metadata) == {"a": 1}
    assert node.metadata is not EMPTY_METADATA

    node.update_metadata({"b": 2})
    assert dict(node.metadata) == {"a": 1, "b": 2}

    other = _leaf("Sheet1", "Z", 9)
    assert other.metadata is EMPTY_METADATA


def test_update_metadata_with_empty_update_keeps_singleton() -> None:
    node = _leaf()
    node.update_metadata({})
    assert node.metadata is EMPTY_METADATA


def test_graph_set_node_metadata_empty_restores_singleton() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf())
    graph.set_node_metadata("Sheet1!A1", {"tag": "x"})
    stored = graph._get_internal_node("Sheet1!A1")
    assert stored is not None
    assert dict(stored.metadata) == {"tag": "x"}

    graph.set_node_metadata("Sheet1!A1", {})
    assert stored.metadata is EMPTY_METADATA


# -------------------------------------------------------------------
# Copies and views must not materialize dicts for the empty case
# -------------------------------------------------------------------


def test_copy_for_projection_shares_singleton_and_copies_dicts() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf())
    tagged = _formula_node()
    tagged.set_metadata({"label": "source"})
    graph.add_node(tagged)
    graph.add_edge("Sheet1!B1", "Sheet1!A1")

    cloned = graph._copy_for_projection()
    assert cloned._nodes["Sheet1!A1"].metadata is EMPTY_METADATA

    cloned_tagged = cloned._nodes["Sheet1!B1"]
    assert dict(cloned_tagged.metadata) == {"label": "source"}
    assert cloned_tagged.metadata is not tagged.metadata
    cloned_tagged.update_metadata({"label": "clone"})
    assert dict(tagged.metadata) == {"label": "source"}


def test_induced_subgraph_shares_singleton() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf())
    graph.add_node(_formula_node())
    graph.add_edge("Sheet1!B1", "Sheet1!A1")

    sub = select_path_induced_subgraph(graph, source_keys=["Sheet1!B1"], target_keys=["Sheet1!A1"])
    assert sub._nodes["Sheet1!A1"].metadata is EMPTY_METADATA
    assert sub._nodes["Sheet1!B1"].metadata is EMPTY_METADATA


def test_node_view_metadata_is_singleton_when_empty_and_read_only_when_set() -> None:
    node = _leaf()
    view = node_to_view(node)
    assert view.metadata is EMPTY_METADATA
    with pytest.raises(TypeError):
        cast(MutableMapping[str, Any], view.metadata)["k"] = 1

    node.set_metadata({"k": "v"})
    view = node_to_view(node)
    assert dict(view.metadata) == {"k": "v"}
    with pytest.raises(TypeError):
        cast(MutableMapping[str, Any], view.metadata)["k"] = "other"


def test_deepcopy_and_pickle_preserve_the_empty_singleton() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf())

    cloned = copy.deepcopy(graph)
    assert cloned._nodes["Sheet1!A1"].metadata is EMPTY_METADATA

    restored: DependencyGraph = pickle.loads(pickle.dumps(graph))
    assert restored._nodes["Sheet1!A1"].metadata is EMPTY_METADATA


# -------------------------------------------------------------------
# Serialization: absent vs. empty metadata round-trip identically
# -------------------------------------------------------------------


def test_json_round_trip_absent_and_empty_metadata_are_identical() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf())
    payload = dependency_graph_to_json(graph)
    assert payload["nodes"][0]["metadata"] == {}

    without_key = copy.deepcopy(payload)
    without_key["nodes"][0].pop("metadata")

    from_empty = dependency_graph_from_json(copy.deepcopy(payload))
    from_absent = dependency_graph_from_json(without_key)

    assert from_empty._nodes["Sheet1!A1"].metadata is EMPTY_METADATA
    assert from_absent._nodes["Sheet1!A1"].metadata is EMPTY_METADATA
    assert dependency_graph_to_json(from_absent) == dependency_graph_to_json(from_empty)


def test_json_round_trip_preserves_populated_metadata() -> None:
    graph = DependencyGraph()
    node = _leaf()
    node.set_metadata({"tag": "x"})
    graph.add_node(node)

    restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    assert dict(restored._nodes["Sheet1!A1"].metadata) == {"tag": "x"}
