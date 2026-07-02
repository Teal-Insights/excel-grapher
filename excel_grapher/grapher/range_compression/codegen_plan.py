"""Codegen evaluation plan over codegen-boundary TACO indexes."""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass

from excel_grapher.core.address_keys import format_cell_key
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey

from .config import TacoBuildConfig
from .index import TacoIndex
from .materialize import materialize_precedents
from .types import CompressedEdge, RangeRef


def range_ref_unit_id(ref: RangeRef) -> str:
    """Return a stable unit id for a compressed dependent range."""
    if ref.min_col == ref.max_col and ref.min_row == ref.max_row:
        return format_cell_key(ref.sheet, ref.min_col, ref.min_row)
    start = format_cell_key(ref.sheet, ref.min_col, ref.min_row)
    return f"{start}:{ref.max_col}{ref.max_row}"


@dataclass(frozen=True, slots=True)
class CompressedUnit:
    """One codegen emission unit for a compressed dependent range."""

    unit_id: str
    edge: CompressedEdge
    dependent: RangeRef


@dataclass(frozen=True, slots=True)
class SingleCellUnit:
    """One codegen emission unit for a single formula cell."""

    unit_id: str
    key: NodeKey


CodegenUnit = CompressedUnit | SingleCellUnit


@dataclass(frozen=True, slots=True)
class CodegenPlan:
    """Evaluation order and coverage map for TACO-aware codegen."""

    compressed_units: tuple[CompressedUnit, ...]
    single_cells: tuple[NodeKey, ...]
    eval_order: tuple[CodegenUnit, ...]
    cell_to_unit: dict[NodeKey, str]
    index: TacoIndex
    config: TacoBuildConfig


def build_codegen_plan(
    graph: DependencyGraph,
    index: TacoIndex,
    config: TacoBuildConfig,
    *,
    closure: Iterable[NodeKey],
) -> CodegenPlan:
    """Build a unit-level evaluation plan from a codegen-boundary TACO index.

    Args:
        graph: Canonical dependency graph.
        index: Codegen-boundary ``TacoIndex``.
        config: Config used to build ``index``.
        closure: Formula cells in the export closure to schedule.
    """
    closure_keys = _formula_cells_in_closure(graph, closure)
    covered_dependents = _covered_dependent_keys(index)
    compressed_units = _compressed_units_from_index(index)
    single_cells = tuple(
        graph.keys(
            order="workbook",
            source=(key for key in closure_keys if key not in covered_dependents),
        )
    )
    cell_to_unit = _build_cell_to_unit(graph, compressed_units, single_cells, config)
    eval_order = _topological_unit_order(
        graph,
        index,
        compressed_units,
        single_cells,
        cell_to_unit,
    )
    return CodegenPlan(
        compressed_units=compressed_units,
        single_cells=single_cells,
        eval_order=eval_order,
        cell_to_unit=cell_to_unit,
        index=index,
        config=config,
    )


def _formula_cells_in_closure(
    graph: DependencyGraph,
    closure: Iterable[NodeKey],
) -> set[NodeKey]:
    keys: set[NodeKey] = set()
    for key in closure:
        node = graph.get_node(key)
        if node is not None and node.formula:
            keys.add(key)
    return keys


def _covered_dependent_keys(index: TacoIndex) -> set[NodeKey]:
    covered: set[NodeKey] = set()
    for edge in index.compressed_edges:
        covered.update(edge.dependent.cell_keys())
    return covered


def _compressed_units_from_index(index: TacoIndex) -> tuple[CompressedUnit, ...]:
    units: list[CompressedUnit] = []
    seen: set[str] = set()
    for edge in index.compressed_edges:
        unit_id = range_ref_unit_id(edge.dependent)
        if unit_id in seen:
            continue
        seen.add(unit_id)
        units.append(
            CompressedUnit(
                unit_id=unit_id,
                edge=edge,
                dependent=edge.dependent,
            )
        )
    return tuple(units)


def _build_cell_to_unit(
    graph: DependencyGraph,
    compressed_units: Sequence[CompressedUnit],
    single_cells: Sequence[NodeKey],
    config: TacoBuildConfig,
) -> dict[NodeKey, str]:
    cell_to_unit: dict[NodeKey, str] = {}
    for key in graph:
        cell_to_unit[key] = key
    for unit in compressed_units:
        for key in unit.dependent.cell_keys():
            cell_to_unit[key] = unit.unit_id
    for key in single_cells:
        cell_to_unit[key] = key
    for key in config.exclude_input_keys:
        cell_to_unit[key] = key
    for key in graph.target_keys():
        cell_to_unit[key] = key
    return cell_to_unit


def _topological_unit_order(
    graph: DependencyGraph,
    index: TacoIndex,
    compressed_units: Sequence[CompressedUnit],
    single_cells: Sequence[NodeKey],
    cell_to_unit: dict[NodeKey, str],
) -> tuple[CodegenUnit, ...]:
    units: list[CodegenUnit] = [
        *compressed_units,
        *(SingleCellUnit(unit_id=key, key=key) for key in single_cells),
    ]
    unit_by_id = {unit.unit_id: unit for unit in units}
    unit_ids = set(unit_by_id)

    deps: dict[str, set[str]] = {uid: set() for uid in unit_ids}
    reverse: dict[str, set[str]] = {uid: set() for uid in unit_ids}

    for unit in units:
        precedent_units = _precedent_units_for_codegen_unit(
            graph,
            index,
            unit,
            cell_to_unit,
            unit_ids,
        )
        for prec_uid in precedent_units:
            if prec_uid == unit.unit_id:
                continue
            deps[unit.unit_id].add(prec_uid)
            reverse[prec_uid].add(unit.unit_id)

    ordered_ids = _kahn_sort(unit_ids, deps, reverse)
    return tuple(unit_by_id[uid] for uid in ordered_ids)


def _precedent_units_for_codegen_unit(
    graph: DependencyGraph,
    index: TacoIndex,
    unit: CodegenUnit,
    cell_to_unit: dict[NodeKey, str],
    unit_ids: set[str],
) -> set[str]:
    precedent_keys: set[NodeKey] = set()
    if isinstance(unit, CompressedUnit):
        for dep_key in unit.dependent.cell_keys():
            precedent_keys.update(materialize_precedents(index, dep_key))
    else:
        precedent_keys.update(graph.get_dependencies(unit.key))

    out: set[str] = set()
    for key in precedent_keys:
        uid = cell_to_unit.get(key, key)
        if uid in unit_ids:
            out.add(uid)
    return out


def _kahn_sort(
    unit_ids: set[str],
    deps: dict[str, set[str]],
    reverse: dict[str, set[str]],
) -> list[str]:
    indegree = {uid: len(deps[uid]) for uid in unit_ids}
    ready = sorted(uid for uid, degree in indegree.items() if degree == 0)
    ordered: list[str] = []

    while ready:
        uid = ready.pop(0)
        ordered.append(uid)
        for dependent in sorted(reverse[uid]):
            indegree[dependent] -= 1
            if indegree[dependent] == 0:
                ready.append(dependent)
                ready.sort()

    if len(ordered) != len(unit_ids):
        remaining = sorted(unit_ids - set(ordered))
        raise ValueError(f"Codegen unit dependency cycle among: {remaining}")
    return ordered
