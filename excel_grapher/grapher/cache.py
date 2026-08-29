from __future__ import annotations

import gzip
import hashlib
import importlib.metadata
import json
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any, Literal, TypedDict, cast

from excel_grapher.core.address_keys import CellKey, parse_node_key
from excel_grapher.core.formula_ast import AstNode
from excel_grapher.core.formula_ast_json import ast_from_json, ast_identity_key, ast_to_json

from .dependency_provenance import DependencyCause, EdgeProvenance
from .graph import DependencyGraph
from .guard import And, CellRef, Compare, GuardExpr, Not, Or, intern_guard
from .guard import Literal as GuardLiteral
from .node import Node, NodeKey

# 6: interned `formula_asts` pool keyed by canonical axis-aware AST JSON.
GRAPH_CACHE_SCHEMA_VERSION = 6


class GraphCacheMeta(TypedDict):
    schema_version: int
    excel_grapher_version: str
    workbook_path: str
    workbook_sha256: str
    workbook_size: int
    workbook_mtime_ns: int
    targets_sha256: str
    extraction_params: dict[str, Any]


class GraphCacheFile(TypedDict):
    meta: GraphCacheMeta
    graph: dict[str, Any]


class CacheValidationPolicy(StrEnum):
    STRICT = "strict"
    PORTABLE = "portable"


def _package_version() -> str:
    try:
        return importlib.metadata.version("excel-grapher")
    except importlib.metadata.PackageNotFoundError:  # pragma: no cover
        return "unknown"


def sha256_file(path: Path) -> str:
    with path.open("rb") as f:
        return hashlib.file_digest(f, "sha256").hexdigest()


def sha256_lines(lines: list[str]) -> str:
    h = hashlib.sha256()
    for line in lines:
        h.update(line.encode("utf-8"))
        h.update(b"\n")
    return h.hexdigest()


def build_graph_cache_meta(
    workbook_path: Path,
    targets: list[str],
    *,
    extraction_params: dict[str, Any] | None = None,
    schema_version: int = GRAPH_CACHE_SCHEMA_VERSION,
) -> GraphCacheMeta:
    resolved = workbook_path.resolve()
    st = resolved.stat()
    return {
        "schema_version": schema_version,
        "excel_grapher_version": _package_version(),
        "workbook_path": str(resolved),
        "workbook_sha256": sha256_file(resolved),
        "workbook_size": int(st.st_size),
        "workbook_mtime_ns": int(st.st_mtime_ns),
        "targets_sha256": sha256_lines(targets),
        "extraction_params": extraction_params or {},
    }


def build_graph_cache_meta_portable(
    targets: list[str],
    *,
    extraction_params: dict[str, Any] | None = None,
    schema_version: int = GRAPH_CACHE_SCHEMA_VERSION,
    excel_grapher_version: str | None = None,
) -> GraphCacheMeta:
    """Build expected cache metadata without access to the workbook file.

    Portable validation enforces only:
    - schema_version
    - excel_grapher_version
    - targets_sha256
    - extraction_params
    """
    return {
        "schema_version": schema_version,
        "excel_grapher_version": excel_grapher_version or _package_version(),
        "workbook_path": "",
        "workbook_sha256": "",
        "workbook_size": 0,
        "workbook_mtime_ns": 0,
        "targets_sha256": sha256_lines(targets),
        "extraction_params": extraction_params or {},
    }


def cache_meta_matches(
    expected: GraphCacheMeta,
    stored: GraphCacheMeta,
    *,
    policy: CacheValidationPolicy = CacheValidationPolicy.STRICT,
) -> bool:
    keys: tuple[str, ...]
    if policy == CacheValidationPolicy.PORTABLE:
        keys = ("schema_version", "excel_grapher_version", "targets_sha256")
    else:
        keys = (
            "schema_version",
            "excel_grapher_version",
            "workbook_path",
            "workbook_sha256",
            "workbook_size",
            "workbook_mtime_ns",
            "targets_sha256",
        )
    for k in keys:
        if stored.get(k) != expected.get(k):
            return False
    return stored.get("extraction_params") == expected.get("extraction_params")


def _json_dump(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    data = json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    if path.suffix.endswith(".gz"):
        with gzip.open(tmp, "wb") as f:
            f.write(data)
    else:
        tmp.write_bytes(data)
    tmp.replace(path)


def _json_load(path: Path) -> object:
    if path.suffix.endswith(".gz"):
        with gzip.open(path, "rb") as f:
            raw = f.read()
    else:
        raw = path.read_bytes()
    return json.loads(raw.decode("utf-8"))


def save_graph_cache(path: Path, graph: DependencyGraph, meta: GraphCacheMeta) -> None:
    payload: GraphCacheFile = {"meta": meta, "graph": dependency_graph_to_json(graph)}
    _json_dump(path, payload)


def try_load_graph_cache(
    path: Path,
    *,
    expected_meta: GraphCacheMeta,
    policy: CacheValidationPolicy = CacheValidationPolicy.STRICT,
) -> DependencyGraph | None:
    if not path.is_file():
        return None
    try:
        payload = _json_load(path)
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(payload, dict):
        return None
    root = cast(dict[str, Any], payload)
    meta = root.get("meta")
    graph = root.get("graph")
    if not isinstance(meta, dict) or not isinstance(graph, dict):
        return None
    stored_meta = cast(GraphCacheMeta, meta)
    if not cache_meta_matches(expected_meta, stored_meta, policy=policy):
        return None
    try:
        return dependency_graph_from_json(graph)
    except (KeyError, TypeError, ValueError):
        return None


class _JsonValue(TypedDict):
    t: Literal["none", "bool", "int", "float", "str", "datetime"]
    v: object


def _value_to_json(v: Any) -> _JsonValue:
    if v is None:
        return {"t": "none", "v": None}
    if isinstance(v, bool):
        return {"t": "bool", "v": v}
    if isinstance(v, int) and not isinstance(v, bool):
        return {"t": "int", "v": v}
    if isinstance(v, float):
        return {"t": "float", "v": v}
    if isinstance(v, str):
        return {"t": "str", "v": v}
    if isinstance(v, datetime):
        return {"t": "datetime", "v": v.isoformat()}
    raise ValueError(f"Unsupported node value type for JSON cache: {type(v)!r}")


def _value_from_json(v: object) -> Any:
    if not isinstance(v, dict):
        raise TypeError("value must be a dict")
    d = cast(dict[str, Any], v)
    t = d.get("t")
    payload = d.get("v")
    if t == "none":
        return None
    if t == "bool":
        if not isinstance(payload, bool):
            raise TypeError("bool payload must be bool")
        return payload
    if t == "int":
        if not isinstance(payload, int) or isinstance(payload, bool):
            raise TypeError("int payload must be int")
        return payload
    if t == "float":
        if not isinstance(payload, (int, float)) or isinstance(payload, bool):
            raise TypeError("float payload must be number")
        return float(payload)
    if t == "str":
        if not isinstance(payload, str):
            raise TypeError("str payload must be str")
        return payload
    if t == "datetime":
        if not isinstance(payload, str):
            raise TypeError("datetime payload must be str")
        return datetime.fromisoformat(payload)
    raise ValueError(f"Unknown JSON value tag: {t!r}")


def _guard_to_json(g: GuardExpr | None) -> object:
    if g is None:
        return None
    if isinstance(g, CellRef):
        return {"type": "cell", "key": g.key}
    if isinstance(g, GuardLiteral):
        return {"type": "lit", "value": _value_to_json(g.value)}
    if isinstance(g, Compare):
        return {
            "type": "cmp",
            "op": g.op,
            "left": _guard_to_json(g.left),
            "right": _guard_to_json(g.right),
        }
    if isinstance(g, Not):
        return {"type": "not", "operand": _guard_to_json(g.operand)}
    if isinstance(g, And):
        return {"type": "and", "operands": [_guard_to_json(o) for o in g.operands]}
    if isinstance(g, Or):
        return {"type": "or", "operands": [_guard_to_json(o) for o in g.operands]}
    raise ValueError(f"Unsupported guard type for JSON cache: {type(g)!r}")


def _guard_from_json(v: object) -> GuardExpr | None:
    if v is None:
        return None
    if not isinstance(v, dict):
        raise TypeError("guard must be dict or null")
    d = cast(dict[str, Any], v)
    typ = d.get("type")
    if typ == "cell":
        key = d.get("key")
        if not isinstance(key, str):
            raise TypeError("cell guard key must be str")
        return intern_guard(CellRef(key=cast(NodeKey, key)))
    if typ == "lit":
        return intern_guard(GuardLiteral(value=_value_from_json(d.get("value"))))
    if typ == "cmp":
        op = d.get("op")
        if not isinstance(op, str):
            raise TypeError("cmp op must be str")
        left = _guard_from_json(d.get("left"))
        right = _guard_from_json(d.get("right"))
        if left is None or right is None:
            raise TypeError("cmp left/right cannot be null")
        return intern_guard(Compare(left=left, op=op, right=right))
    if typ == "not":
        operand = _guard_from_json(d.get("operand"))
        if operand is None:
            raise TypeError("not operand cannot be null")
        return intern_guard(Not(operand=operand))
    if typ == "and":
        ops = d.get("operands")
        if not isinstance(ops, list):
            raise TypeError("and operands must be list")
        parts: list[GuardExpr] = []
        for o in ops:
            g = _guard_from_json(o)
            if g is None:
                raise TypeError("and operands cannot be null")
            parts.append(g)
        return intern_guard(And(tuple(parts)))
    if typ == "or":
        ops = d.get("operands")
        if not isinstance(ops, list):
            raise TypeError("or operands must be list")
        parts_o: list[GuardExpr] = []
        for o in ops:
            g = _guard_from_json(o)
            if g is None:
                raise TypeError("or operands cannot be null")
            parts_o.append(g)
        return intern_guard(Or(tuple(parts_o)))
    raise ValueError(f"Unknown guard JSON type: {typ!r}")


def _edge_provenance_to_json(p: EdgeProvenance) -> dict[str, Any]:
    return {
        "causes": sorted(c.name for c in p.causes),
        "direct_sites_normalized": [list(x) for x in p.direct_sites_normalized],
    }


def _edge_provenance_from_json(v: object) -> EdgeProvenance:
    if not isinstance(v, dict):
        raise TypeError("provenance must be dict")
    d = cast(dict[str, Any], v)
    causes_v = d.get("causes")
    if not isinstance(causes_v, list) or not all(isinstance(x, str) for x in causes_v):
        raise TypeError("provenance.causes must be list[str]")
    causes = DependencyCause(0)
    for name in causes_v:
        try:
            causes |= DependencyCause[name]
        except KeyError as e:
            raise TypeError(f"unknown provenance cause: {name!r}") from e
    # Schema 3 payloads also carry `direct_sites_formula`; it is ignored.
    dsn = d.get("direct_sites_normalized", [])
    if not isinstance(dsn, list):
        raise TypeError("provenance sites must be lists")
    return EdgeProvenance(
        causes=causes,
        direct_sites_normalized=tuple((int(a), int(b)) for a, b in dsn),
    )


def _formula_ast_intern_key(ast: AstNode) -> str:
    return ast_identity_key(ast)


def _formula_asts_to_json(graph: DependencyGraph) -> dict[str, Any]:
    pool: dict[str, Any] = {}
    for _key, node in graph.formula_nodes():
        if node.formula_ast is None:
            continue
        intern_key = _formula_ast_intern_key(node.formula_ast)
        if intern_key in pool:
            continue
        pool[intern_key] = ast_to_json(node.formula_ast)
    return pool


def _formula_asts_from_json(payload: dict[str, Any]) -> dict[str, AstNode]:
    raw = payload.get("formula_asts", {})
    if not isinstance(raw, dict):
        raise TypeError("formula_asts must be an object")
    pool: dict[str, AstNode] = {}
    for key, encoded in raw.items():
        if not isinstance(key, str):
            raise TypeError("formula_asts keys must be strings")
        pool[key] = ast_from_json(encoded)
    return pool


def dependency_graph_to_json(graph: DependencyGraph) -> dict[str, Any]:
    nodes: list[dict[str, Any]] = []
    for key in graph:
        node = graph.get_node(key)
        if node is None:
            continue
        payload = {
            "key": key,
            "address": key,
            "sheet": node.sheet,
            "column": node.column,
            "row": node.row,
            "min_col": node.min_col,
            "min_row": node.min_row,
            "max_col": node.max_col,
            "max_row": node.max_row,
            "formula": node.formula,
            "normalized_formula": node.normalized_formula,
            "value": _value_to_json(node.value),
            "is_leaf": node.is_leaf,
            "is_target": node.is_target,
            "metadata": dict(node.metadata),
        }
        if node.formula_ast is not None:
            payload["formula_ast_key"] = _formula_ast_intern_key(node.formula_ast)
        nodes.append(payload)

    edges: list[dict[str, Any]] = []
    for from_key in graph:
        for to_key in graph.get_dependencies(from_key):
            edge = graph.get_edge_attrs(from_key, to_key)
            attrs: dict[str, Any] = {}
            if edge.provenance is not None:
                attrs["provenance"] = _edge_provenance_to_json(edge.provenance)
            edges.append(
                {
                    "from": from_key,
                    "to": to_key,
                    "guard": _guard_to_json(edge.guard),
                    "attrs": attrs,
                }
            )

    return {
        "nodes": nodes,
        "edges": edges,
        "leaf_classification": graph.leaf_classification,
        "formula_asts": _formula_asts_to_json(graph),
    }


def dependency_graph_from_json(payload: dict[str, Any]) -> DependencyGraph:
    nodes_v = payload["nodes"]
    edges_v = payload["edges"]
    leaf_cls = payload.get("leaf_classification")
    if not isinstance(nodes_v, list) or not isinstance(edges_v, list):
        raise TypeError("graph nodes/edges must be lists")

    g = DependencyGraph()
    if leaf_cls is not None:
        if not isinstance(leaf_cls, dict) or not all(
            isinstance(k, str) and isinstance(v, str) for k, v in leaf_cls.items()
        ):
            raise TypeError("leaf_classification must be dict[str,str]")
        g.leaf_classification = cast(dict[str, str], leaf_cls)

    formula_asts = _formula_asts_from_json(payload)

    for n in nodes_v:
        if not isinstance(n, dict):
            raise TypeError("node must be dict")
        address = n.get("address")
        if address is None:
            address = n.get("key")
        formula = cast(str | None, n["formula"])
        normalized_formula = cast(str | None, n["normalized_formula"])
        intern_key = n.get("formula_ast_key")
        if not isinstance(intern_key, str):
            intern_key = None
        formula_ast = formula_asts.get(intern_key) if intern_key is not None else None
        value = _value_from_json(n["value"])
        is_leaf = bool(n["is_leaf"])
        is_target = bool(n.get("is_target", False))
        metadata = cast(dict[str, Any], n.get("metadata", {}))
        if address is not None:
            parsed = parse_node_key(cast(str, address))
            if not isinstance(parsed, CellKey):
                raise TypeError(
                    f"Graph cache nodes must be cells; got {type(parsed).__name__}: {address!r}"
                )
            node = Node(
                sheet=cast(str | None, n.get("sheet")),
                column=cast(str | None, n.get("column")),
                row=None if n.get("row") is None else int(n["row"]),
                formula=formula,
                normalized_formula=normalized_formula,
                value=value,
                is_leaf=is_leaf,
                is_target=is_target,
                metadata=metadata,
                address=parsed,
                formula_ast=formula_ast,
            )
        else:
            node = Node(
                sheet=cast(str, n["sheet"]),
                column=cast(str, n["column"]),
                row=int(n["row"]),
                formula=formula,
                normalized_formula=normalized_formula,
                value=value,
                is_leaf=is_leaf,
                is_target=is_target,
                metadata=metadata,
                formula_ast=formula_ast,
            )
        g.add_node(node)

    for e in edges_v:
        if not isinstance(e, dict):
            raise TypeError("edge must be dict")
        from_key = cast(NodeKey, e["from"])
        to_key = cast(NodeKey, e["to"])
        guard = _guard_from_json(e.get("guard"))
        attrs = e.get("attrs", {})
        if not isinstance(attrs, dict):
            raise TypeError("edge attrs must be dict")
        provenance = None
        if "provenance" in attrs:
            provenance = _edge_provenance_from_json(attrs["provenance"])
        g.add_edge(from_key, to_key, guard=guard, provenance=provenance)

    return g
