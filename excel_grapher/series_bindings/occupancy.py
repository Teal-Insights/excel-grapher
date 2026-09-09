"""Cell occupancy pairing for mixed series (issue #708).

An `input` or `constant` series may name a retained graph-leaf cell that an
`internal`/`output` series already owns. Formula cells stay single-owner.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import Any, Literal

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.graph_predicates import (
    is_graph_formula_node,
    is_graph_leaf,
)
from excel_grapher.series_bindings.normalize import (
    effective_validation,
    has_constant_direction,
    has_input_direction,
    has_internal_direction,
    has_output_direction,
)

Direction = Literal["input", "constant", "internal", "output"]

_FORMULA_DIRECTIONS = frozenset({"internal", "output"})
_CLAIMANT_DIRECTIONS = frozenset({"input", "constant"})


def binding_direction(series: dict[str, Any]) -> Direction | None:
    """Return the series direction, or `None` when no direction is declared."""
    if has_output_direction(series):
        return "output"
    if has_internal_direction(series):
        return "internal"
    if has_input_direction(series):
        return "input"
    if has_constant_direction(series):
        return "constant"
    return None


def occupancy_addresses(
    graph: DependencyGraph,
    series: dict[str, Any],
    addresses: Sequence[str],
) -> list[str]:
    """Return cells this series contributes to occupancy checks.

    Matches inverted-tree catalog filtering: input/constant ranges stay whole;
    formula series keep graph formulas and on-graph leaves for `layout: series`,
    and the full rectangle for `layout: matrix`.
    """
    direction = binding_direction(series)
    if direction not in _FORMULA_DIRECTIONS:
        return list(addresses)
    layout = str(series.get("layout") or "scalar")
    if layout == "row_series":
        layout = "series"
    validation = effective_validation(series)
    if direction == "internal" and not validation.get("intersect_graph_formulas", True):
        return list(addresses)
    if direction == "output" and not validation.get("intersect_graph_nodes", True):
        return list(addresses)
    if layout == "matrix":
        return list(addresses)
    return [
        address
        for address in addresses
        if is_graph_formula_node(graph, address) or is_graph_leaf(graph, address)
    ]


@dataclass(frozen=True, slots=True)
class BoundLeafPair:
    """Allowed mixed-series occupancy: formula owner plus leaf claimant."""

    owner_id: str
    claimant_id: str
    address: str


class CellOccupancyError(ValueError):
    """Two or more series claim a cell that is not an allowed bound-leaf pair."""

    def __init__(self, address: str, first_id: str, second_id: str) -> None:
        self.address = address
        self.first_id = first_id
        self.second_id = second_id
        super().__init__(f"cell {address} is bound to both {first_id!r} and {second_id!r}")


def resolve_occupants(
    occupants: Sequence[tuple[str, str]],
    *,
    address: str,
    is_graph_leaf: bool,
) -> str | BoundLeafPair:
    """Return the unique owner id, or an allowed bound-leaf pairing.

    `occupants` is `(series_id, direction)` in bindings order.

    Raises:
        CellOccupancyError: Overlap is not a single graph-leaf pairing.
    """
    unique: list[tuple[str, str]] = []
    seen: set[str] = set()
    for series_id, direction in occupants:
        if series_id in seen:
            continue
        seen.add(series_id)
        unique.append((series_id, direction))
    if len(unique) == 1:
        return unique[0][0]
    if len(unique) != 2:
        raise CellOccupancyError(address, unique[0][0], unique[1][0])
    (id_a, dir_a), (id_b, dir_b) = unique
    if dir_a in _FORMULA_DIRECTIONS and dir_b in _CLAIMANT_DIRECTIONS and is_graph_leaf:
        return BoundLeafPair(owner_id=id_a, claimant_id=id_b, address=address)
    if dir_b in _FORMULA_DIRECTIONS and dir_a in _CLAIMANT_DIRECTIONS and is_graph_leaf:
        return BoundLeafPair(owner_id=id_b, claimant_id=id_a, address=address)
    raise CellOccupancyError(address, id_a, id_b)
