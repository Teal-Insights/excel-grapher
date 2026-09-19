"""Low-peak pickle encoding for `DependencyGraph`.

CPython's unpickler memo retains every reconstructed object until `pickle.load`
returns. A single state-dict pickle would therefore peak near 2x final size:
indexed adjacency sets sit in the memo while live string-keyed maps are built
beside them.

This module writes a two-frame payload (nodes, then CSR+CSC arrays). Each
frame's unpickler is discarded before the next frame is read, so peak stays
close to final resident size. `DependencyGraph.__reduce_ex__` wraps a
gzip-compressed multipart blob so `pickle.loads` also stays near final size;
prefer `dump_graph` / `load_graph` for files (no outer bytes envelope).
"""

from __future__ import annotations

import array
import gzip
import io
import pickle
from pathlib import Path
from typing import Any, BinaryIO, cast

from .csr_adjacency import build_csr_csc
from .dependency_provenance import EdgeProvenance
from .node import Node, NodeKey

# Logical payload: magic + little-endian version + two pickle frames.
# `dumps_graph_blob` gzip-compresses that payload for the pickle reduce path.
_GRAPH_BLOB_MAGIC = b"EGDG"
_GRAPH_BLOB_VERSION = 6
_GRAPH_BLOB_MIN_VERSION = 5
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
            _require_supported_version(version)
            return _read_graph_frames(binary, version=version)
        binary.seek(0)
        return pickle.load(binary)


def _require_supported_version(version: int) -> None:
    if version < _GRAPH_BLOB_MIN_VERSION or version > _GRAPH_BLOB_VERSION:
        raise TypeError("Unsupported or corrupted DependencyGraph pickle; rebuild the graph cache.")


def _load_header_and_frames(buf: BinaryIO) -> Any:
    header = buf.read(len(_GRAPH_BLOB_HEADER))
    if not header.startswith(_GRAPH_BLOB_MAGIC):
        raise TypeError("Unsupported or corrupted DependencyGraph pickle; rebuild the graph cache.")
    version = int.from_bytes(header[4:8], "little")
    _require_supported_version(version)
    return _read_graph_frames(buf, version=version)


def _csr_payload(
    graph: Any,
) -> tuple[list[NodeKey], array.array[int], array.array[int], array.array[int], array.array[int]]:
    """Return `(csr_keys, row_ptr, col_idx, col_ptr, row_idx)` without mutating `graph`."""
    if not graph._staging:
        return (
            list(graph._csr_keys),
            graph._row_ptr,
            graph._col_idx,
            graph._col_ptr,
            graph._row_idx,
        )
    empty: tuple[NodeKey, ...] = ()
    edges = graph._edges
    built = build_csr_csc(
        tuple(graph._nodes),
        lambda key: edges.get(key, empty),
        is_node=graph._nodes.__contains__,
    )
    return built.keys, built.row_ptr, built.col_idx, built.col_ptr, built.row_idx


def _write_graph_frames(graph: Any, buf: BinaryIO) -> None:
    from .graph import _collect_graph_keys

    keys_sorted = _collect_graph_keys(graph)
    idx = {k: i for i, k in enumerate(keys_sorted)}
    csr_keys, row_ptr, col_idx, col_ptr, row_idx = _csr_payload(graph)
    node_keys = csr_keys
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

    pickle.dump(
        {
            "row_ptr": row_ptr,
            "col_idx": col_idx,
            "col_ptr": col_ptr,
            "row_idx": row_idx,
            "_guards": [(idx[a], idx[b], g) for (a, b), g in graph._guards.items()],
            "_edge_provenance": [
                (idx[a], idx[b], p) for (a, b), p in graph._edge_provenance.items()
            ],
        },
        buf,
        protocol=pickle.HIGHEST_PROTOCOL,
    )


def _read_graph_frames(buf: BinaryIO, *, version: int) -> Any:
    from .graph import DependencyGraph, _intern_guard_cell_refs

    part1 = pickle.load(buf)
    keys: list[str] = part1["keys"]
    node_keys: list[str] = part1["node_keys"]
    nodes: list[Node] = part1["nodes"]
    key_index = {s: i for i, s in enumerate(keys)}

    graph = DependencyGraph.__new__(DependencyGraph)
    interned_nodes = {keys[key_index[k]]: n for k, n in zip(node_keys, nodes, strict=True)}
    graph._nodes = interned_nodes
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
    graph.cell_type_env = None
    graph._value_generation = 0
    graph._edges = {}
    graph._reverse_edges = {}
    graph._staging = True
    graph._csr_keys = []
    graph._node_index = {}
    graph._row_ptr = array.array("I")
    graph._col_idx = array.array("I")
    graph._col_ptr = array.array("I")
    graph._row_idx = array.array("I")
    del part1, nodes

    part2 = pickle.load(buf)
    graph._guards = {
        (keys[a], keys[b]): _intern_guard_cell_refs(g, keys, key_index=key_index)
        for a, b, g in part2["_guards"]
    }
    graph._edge_provenance = {
        (keys[a], keys[b]): cast(EdgeProvenance, p) for a, b, p in part2["_edge_provenance"]
    }
    if version >= 6:
        csr_keys = [keys[key_index[k]] for k in node_keys]
        graph._csr_keys = csr_keys
        graph._node_index = {k: i for i, k in enumerate(csr_keys)}
        graph._row_ptr = part2["row_ptr"]
        graph._col_idx = part2["col_idx"]
        graph._col_ptr = part2["col_ptr"]
        graph._row_idx = part2["row_idx"]
        graph._staging = False
        del part2
        return graph

    edge_src: array.array[int] = part2["edge_src"]
    edge_dst: array.array[int] = part2["edge_dst"]
    for s, d in zip(edge_src, edge_dst, strict=True):
        src = keys[s]
        dst = keys[d]
        graph._edges.setdefault(src, set()).add(dst)
        graph._reverse_edges.setdefault(dst, set()).add(src)
    del part2, edge_src, edge_dst
    graph.rebuild_adjacency()
    return graph
