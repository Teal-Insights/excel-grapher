"""Load solver-oriented may-cycle MCVE fragments (`kind: may_cycle_solver_mcve`)."""

from __future__ import annotations

import gzip
import json
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_grapher.core.address_keys import CellKey, parse_node_key
from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    EnumDomain,
    IntervalDomain,
    RealIntervalDomain,
    normalize_cell_type_env_key,
)

from .cache import _guard_from_json
from .graph import DependencyGraph
from .guard import GuardExpr
from .node import Node, NodeKey


@dataclass(frozen=True, slots=True)
class SolverMcveNode:
    """One cell in a solver MCVE fragment."""

    in_scc: bool
    in_guard_cone: bool
    is_leaf: bool
    normalized_formula: str | None
    value: Any


@dataclass(frozen=True, slots=True)
class SolverMcveEdge:
    """Directed dependency edge (`from` depends on `to`)."""

    src: NodeKey
    dst: NodeKey
    guard: GuardExpr | None


@dataclass(frozen=True, slots=True)
class SolverMcve:
    """In-memory `may_cycle_solver_mcve` fragment (schema 1.0.0)."""

    kind: str
    schema_version: str
    scc_index: int
    seed: NodeKey | None
    scc_members: tuple[NodeKey, ...]
    guard_refs: tuple[NodeKey, ...]
    guard_cone_outside_scc: tuple[NodeKey, ...]
    feasible_cycle_path: tuple[NodeKey, ...]
    missing_leaf_constraints: tuple[NodeKey, ...]
    leaf_constraints: Mapping[NodeKey, Mapping[str, Any]]
    nodes: Mapping[NodeKey, SolverMcveNode]
    intra_scc_edges: tuple[SolverMcveEdge, ...]
    guard_cone_edges: tuple[SolverMcveEdge, ...]

    def cell_type_env(self) -> dict[str, CellType]:
        """Build a `CellTypeEnv` from serialized `leaf_constraints`."""
        return leaf_constraints_to_cell_type_env(self.leaf_constraints)

    def to_graph(self, *, include_guard_cone: bool = True) -> DependencyGraph:
        """Materialize a `DependencyGraph` for cycle feasibility replay.

        Intra-SCC edges are always attached. Guard-cone nodes (and their
        identity formulas) are included so alias rewriting can run. Cone edges
        are omitted: they are outside the SCC and do not participate in
        `cycle_report`.
        """
        graph = DependencyGraph()
        keys = set(self.scc_members)
        if include_guard_cone:
            keys.update(self.guard_cone_outside_scc)
            keys.update(self.nodes)
        for key in keys:
            rec = self.nodes.get(key)
            parsed = parse_node_key(str(key))
            if not isinstance(parsed, CellKey):
                continue
            formula = None
            if rec is not None and rec.in_guard_cone:
                formula = rec.normalized_formula
            graph.add_node(
                Node(
                    address=parsed,
                    normalized_formula=formula,
                    value=None if rec is None else rec.value,
                    is_leaf=True if rec is None else rec.is_leaf,
                )
            )
        for edge in self.intra_scc_edges:
            graph.add_edge(edge.src, edge.dst, guard=edge.guard)
        graph.cell_type_env = self.cell_type_env()
        return graph


def load_solver_mcve(path: str | Path) -> SolverMcve:
    """Load a solver MCVE from JSON or gzip-compressed JSON."""
    dest = Path(path)
    if dest.suffix == ".gz" or dest.name.endswith(".json.gz"):
        with gzip.open(dest, "rt", encoding="utf-8") as handle:
            payload = json.load(handle)
    else:
        payload = json.loads(dest.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise TypeError("solver MCVE root must be an object")
    return solver_mcve_from_mapping(payload)


def solver_mcve_from_mapping(payload: Mapping[str, Any]) -> SolverMcve:
    """Parse a solver MCVE mapping (already-decoded JSON)."""
    kind = payload.get("kind")
    if kind != "may_cycle_solver_mcve":
        raise ValueError(f"expected kind 'may_cycle_solver_mcve', got {kind!r}")
    schema = payload.get("schema_version", "1.0.0")
    if not isinstance(schema, str):
        raise TypeError("schema_version must be str")

    nodes_raw = payload.get("nodes")
    if not isinstance(nodes_raw, dict):
        raise TypeError("nodes must be an object")
    nodes: dict[NodeKey, SolverMcveNode] = {}
    for key, rec in nodes_raw.items():
        if not isinstance(key, str) or not isinstance(rec, dict):
            raise TypeError("nodes entries must be string keys to objects")
        nodes[key] = SolverMcveNode(
            in_scc=bool(rec.get("in_scc")),
            in_guard_cone=bool(rec.get("in_guard_cone")),
            is_leaf=bool(rec.get("is_leaf")),
            normalized_formula=rec.get("normalized_formula"),
            value=rec.get("value"),
        )

    return SolverMcve(
        kind=kind,
        schema_version=schema,
        scc_index=int(payload.get("scc_index", 0)),
        seed=payload.get("seed"),
        scc_members=_str_tuple(payload.get("scc_members")),
        guard_refs=_str_tuple(payload.get("guard_refs")),
        guard_cone_outside_scc=_str_tuple(payload.get("guard_cone_outside_scc")),
        feasible_cycle_path=_str_tuple(payload.get("feasible_cycle_path")),
        missing_leaf_constraints=_str_tuple(payload.get("missing_leaf_constraints")),
        leaf_constraints=_mapping(payload.get("leaf_constraints"), "leaf_constraints"),
        nodes=nodes,
        intra_scc_edges=_edges(payload.get("intra_scc_edges"), "intra_scc_edges"),
        guard_cone_edges=_edges(payload.get("guard_cone_edges"), "guard_cone_edges"),
    )


def leaf_constraints_to_cell_type_env(
    leaf_constraints: Mapping[str, Any],
) -> dict[str, CellType]:
    """Convert solver-MCVE `leaf_constraints` objects to a `CellTypeEnv`."""
    env: dict[str, CellType] = {}
    for key, rec in leaf_constraints.items():
        if not isinstance(rec, dict):
            raise TypeError(f"leaf constraint for {key} must be an object")
        constraint = rec.get("constraint")
        if not isinstance(constraint, dict):
            raise TypeError(f"leaf constraint.constraint for {key} must be an object")
        env[normalize_cell_type_env_key(key)] = _cell_type_from_mcve_constraint(constraint)
    return env


def _cell_type_from_mcve_constraint(constraint: Mapping[str, Any]) -> CellType:
    kind = constraint.get("kind")
    if kind == "literal":
        values = constraint.get("values")
        if not isinstance(values, list):
            raise TypeError("literal constraint values must be a list")
        literal_values = tuple(values)
        return CellType(
            kind=_kind_from_values(literal_values),
            enum=EnumDomain(values=frozenset(literal_values)),
        )
    if kind == "annotated":
        int_domain: IntervalDomain | None = None
        real_domain: RealIntervalDomain | None = None
        for meta in constraint.get("meta") or []:
            if not isinstance(meta, dict):
                continue
            meta_type = meta.get("type")
            if meta_type == "Between":
                int_domain = IntervalDomain(min=meta.get("min"), max=meta.get("max"))
            elif meta_type == "RealBetween":
                lo, hi = meta.get("min"), meta.get("max")
                real_domain = RealIntervalDomain(
                    min=None if lo is None else float(lo),
                    max=None if hi is None else float(hi),
                )
        return CellType(
            kind=CellKind.NUMBER,
            interval=int_domain,
            real_interval=real_domain,
        )
    raise ValueError(f"unsupported leaf constraint kind: {kind!r}")


def _kind_from_values(values: tuple[object, ...]) -> CellKind:
    if values and all(isinstance(v, bool) for v in values):
        return CellKind.BOOL
    if values and all(isinstance(v, int) and not isinstance(v, bool) for v in values):
        return CellKind.NUMBER
    if values and all(isinstance(v, str) for v in values):
        return CellKind.STRING
    return CellKind.ANY


def _str_tuple(value: object) -> tuple[NodeKey, ...]:
    if value is None:
        return ()
    if not isinstance(value, list):
        raise TypeError("expected a list of strings")
    out: list[NodeKey] = []
    for item in value:
        if not isinstance(item, str):
            raise TypeError("expected string entries")
        out.append(item)
    return tuple(out)


def _mapping(value: object, name: str) -> dict[str, Any]:
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise TypeError(f"{name} must be an object")
    return {str(k): v for k, v in value.items()}


def _edges(value: object, name: str) -> tuple[SolverMcveEdge, ...]:
    if value is None:
        return ()
    if not isinstance(value, list):
        raise TypeError(f"{name} must be a list")
    out: list[SolverMcveEdge] = []
    for item in value:
        if not isinstance(item, dict):
            raise TypeError(f"{name} entries must be objects")
        src = item.get("from")
        dst = item.get("to")
        if not isinstance(src, str) or not isinstance(dst, str):
            raise TypeError(f"{name} from/to must be strings")
        out.append(SolverMcveEdge(src=src, dst=dst, guard=_guard_from_json(item.get("guard"))))
    return tuple(out)
