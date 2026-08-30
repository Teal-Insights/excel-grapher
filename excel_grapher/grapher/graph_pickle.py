"""Low-peak pickle encoding for `DependencyGraph`.

CPython's unpickler memo retains every reconstructed object until `pickle.load`
returns. The legacy `__getstate__` format therefore peaks near 2x final size:
indexed adjacency sets sit in the memo while `__setstate__` builds the live
string-keyed maps beside them.

This module writes a two-frame payload (nodes, then COO edges). Each frame's
unpickler is discarded before the next frame is read, so peak stays close to
final resident size. `DependencyGraph.__reduce_ex__` wraps a gzip-compressed
multipart blob so `pickle.loads` also stays near final size; prefer
`dump_graph` / `load_graph` for files (no outer bytes envelope).
"""

from __future__ import annotations

import array
import gzip
import io
import pickle
from pathlib import Path
from typing import Any, BinaryIO, cast

from .dependency_provenance import EdgeProvenance
from .node import Node, NodeKey

# Logical payload: magic + little-endian version + two pickle frames.
# `dumps_graph_blob` gzip-compresses that payload for the pickle reduce path.
_GRAPH_BLOB_MAGIC = b"EGDG"
_GRAPH_BLOB_VERSION = 5
_GRAPH_BLOB_HEADER = _GRAPH_BLOB_MAGIC + _GRAPH_BLOB_VERSION.to_bytes(4, "little")
_GZIP_MAGIC = b"\x1f\x8b"


def dumps_graph_blob(graph: Any) -> bytes:
    """Serialize `graph` to a gzip-compressed multipart pickle blob."""
    buf = io.BytesIO()
    with gzip.GzipFile(fileobj=buf, mode="wb", compresslevel=1) as handle:
        binary = cast(BinaryIO, handle)
        binary.write(_GRAPH_BLOB_HEADER)
        _write_graph_frames(graph, binary)
    return buf.getvalue()


def loads_graph_blob(blob: bytes) -> Any:
    """Restore a graph from `dumps_graph_blob` / `__reduce_ex__` payload."""
    if blob.startswith(_GZIP_MAGIC):
        with gzip.GzipFile(fileobj=io.BytesIO(blob), mode="rb") as handle:
            return _load_header_and_frames(cast(BinaryIO, handle))
    return _load_header_and_frames(io.BytesIO(blob))


def dump_graph(
    graph: Any,
    path: str | Path,
    *,
    compress: bool | None = None,
) -> None:
    """Write `graph` to `path` using the low-peak multipart pickle format.

    Args:
        graph: Graph to serialize.
        path: Destination path. Parent directories are created as needed.
        compress: Gzip the file when True. When omitted, gzip is used if `path`
            ends with `.gz`.
    """
    dest = Path(path)
    use_gzip = dest.suffix == ".gz" if compress is None else compress
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".tmp")
    opener = gzip.open if use_gzip else open
    with opener(tmp, "wb") as handle:
        binary = cast(BinaryIO, handle)
        binary.write(_GRAPH_BLOB_HEADER)
        _write_graph_frames(graph, binary)
    tmp.replace(dest)


def load_graph(path: str | Path) -> Any:
    """Load a graph from `dump_graph` output or a legacy `pickle` stream.

    Sniffs the `EGDG` multipart header first. If absent, falls back to
    `pickle.load` so older single-object pickles still open.

    `formula_shapes` and `preparsed_formulas` are omitted (`None`). Call
    `warm_formula_shapes` / `warm_preparsed_formulas` after load if you want
    those overlays. A live `FormulaEvaluator` does not refresh compiled
    shape helpers from a later rewarm; construct a new evaluator.
    """
    source = Path(path)
    opener = gzip.open if source.suffix == ".gz" else open
    with opener(source, "rb") as handle:
        binary = cast(BinaryIO, handle)
        header = binary.read(len(_GRAPH_BLOB_HEADER))
        if header.startswith(_GRAPH_BLOB_MAGIC):
            version = int.from_bytes(header[4:8], "little")
            if version != _GRAPH_BLOB_VERSION:
                raise TypeError(
                    "Unsupported or corrupted DependencyGraph pickle; rebuild the graph cache."
                )
            return _read_graph_frames(binary)
        binary.seek(0)
        return pickle.load(binary)


def _load_header_and_frames(buf: BinaryIO) -> Any:
    header = buf.read(len(_GRAPH_BLOB_HEADER))
    if not header.startswith(_GRAPH_BLOB_MAGIC):
        raise TypeError("Unsupported or corrupted DependencyGraph pickle; rebuild the graph cache.")
    version = int.from_bytes(header[4:8], "little")
    if version != _GRAPH_BLOB_VERSION:
        raise TypeError("Unsupported or corrupted DependencyGraph pickle; rebuild the graph cache.")
    return _read_graph_frames(buf)


def _write_graph_frames(graph: Any, buf: BinaryIO) -> None:
    from .graph import _collect_graph_keys

    keys_sorted = _collect_graph_keys(graph)
    idx = {k: i for i, k in enumerate(keys_sorted)}
    node_keys = [k for k in keys_sorted if k in graph._nodes]
    nodes = [graph._nodes[k] for k in node_keys]

    # Frame 1: nodes + graph-level metadata (no adjacency).
    pickle.dump(
        {
            "keys": keys_sorted,
            "node_keys": node_keys,
            "nodes": nodes,
            "_hooks": graph._hooks,
            "leaf_classification": graph.leaf_classification,
            "sheet_order": list(graph.sheet_order) if graph.sheet_order is not None else None,
            "named_ranges": dict(graph.named_ranges) if graph.named_ranges else None,
            "named_range_ranges": (
                dict(graph.named_range_ranges) if graph.named_range_ranges else None
            ),
        },
        buf,
        protocol=pickle.HIGHEST_PROTOCOL,
    )

    edge_src, edge_dst = _edges_to_coo(graph._edges, idx)
    pickle.dump(
        {
            "edge_src": edge_src,
            "edge_dst": edge_dst,
            "_guards": [(idx[a], idx[b], g) for (a, b), g in graph._guards.items()],
            "_edge_provenance": [
                (idx[a], idx[b], p) for (a, b), p in graph._edge_provenance.items()
            ],
        },
        buf,
        protocol=pickle.HIGHEST_PROTOCOL,
    )


def _read_graph_frames(buf: BinaryIO) -> Any:
    from .graph import DependencyGraph, _intern_guard_cell_refs

    part1 = pickle.load(buf)
    keys: list[str] = part1["keys"]
    node_keys: list[str] = part1["node_keys"]
    nodes: list[Node] = part1["nodes"]
    key_index = {s: i for i, s in enumerate(keys)}

    graph = DependencyGraph.__new__(DependencyGraph)
    # Intern node-map keys against the shared `keys` list.
    graph._nodes = {keys[key_index[k]]: n for k, n in zip(node_keys, nodes, strict=True)}
    graph._hooks = part1["_hooks"]
    lc = part1["leaf_classification"]
    if lc:
        graph.leaf_classification = {keys[key_index[k]]: v for k, v in lc.items()}
    else:
        graph.leaf_classification = None
    sheet_order = part1.get("sheet_order")
    graph.sheet_order = list(sheet_order) if sheet_order else None
    nr = part1.get("named_ranges")
    graph.named_ranges = dict(nr) if nr else None
    nrr = part1.get("named_range_ranges")
    graph.named_range_ranges = dict(nrr) if nrr else None
    graph.sheet_bounds = None
    graph.preparsed_formulas = None
    graph.formula_shapes = None
    del part1, nodes, node_keys

    part2 = pickle.load(buf)
    edge_src: array.array[int] = part2["edge_src"]
    edge_dst: array.array[int] = part2["edge_dst"]
    graph._edges = {k: set() for k in graph._nodes}
    graph._reverse_edges = {k: set() for k in graph._nodes}
    for s, d in zip(edge_src, edge_dst, strict=True):
        src = keys[s]
        dst = keys[d]
        graph._edges.setdefault(src, set()).add(dst)
        graph._reverse_edges.setdefault(dst, set()).add(src)
    graph._guards = {
        (keys[a], keys[b]): _intern_guard_cell_refs(g, keys, key_index=key_index)
        for a, b, g in part2["_guards"]
    }
    graph._edge_provenance = {
        (keys[a], keys[b]): cast(EdgeProvenance, p) for a, b, p in part2["_edge_provenance"]
    }
    del part2, edge_src, edge_dst
    return graph


def _edges_to_coo(
    edges: dict[NodeKey, set[NodeKey]],
    idx: dict[str, int],
) -> tuple[array.array[int], array.array[int]]:
    src = array.array("I")
    dst = array.array("I")
    for key, deps in edges.items():
        s = idx[key]
        for dep in deps:
            src.append(s)
            dst.append(idx[dep])
    return src, dst
