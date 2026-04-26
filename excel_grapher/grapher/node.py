from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum, auto
from types import MappingProxyType
from typing import Any

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_cell_key as _format_node_key

NodeKey = str  # Always in the form "SheetName!A1" or "'Sheet Name'!A1" for quoted sheets


class ValueType(Enum):
    NUMBER = auto()
    STRING = auto()
    BOOLEAN = auto()
    ERROR = auto()
    DATETIME = auto()
    EMPTY = auto()
    UNKNOWN = auto()


def _classify_value(value: Any, *, is_leaf: bool) -> ValueType:
    if value is None:
        return ValueType.EMPTY if is_leaf else ValueType.UNKNOWN
    # bool must be checked before int, since bool subclasses int.
    if isinstance(value, bool):
        return ValueType.BOOLEAN
    if isinstance(value, (int, float)):
        return ValueType.NUMBER
    if isinstance(value, datetime):
        return ValueType.DATETIME
    if isinstance(value, str):
        if value.startswith("#"):
            return ValueType.ERROR
        return ValueType.STRING
    return ValueType.UNKNOWN


@dataclass
class Node:
    sheet: str
    column: str
    row: int
    formula: str | None
    normalized_formula: str | None
    value: Any
    is_leaf: bool
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def key(self) -> NodeKey:
        return _format_node_key(self.sheet, self.column, self.row)

    @property
    def address(self) -> str:
        return f"{self.column}{self.row}"

    @property
    def column_index(self) -> int:
        return int(fastpyxl.utils.cell.column_index_from_string(self.column))

    @property
    def value_type(self) -> ValueType:
        return _classify_value(self.value, is_leaf=self.is_leaf)


@dataclass(frozen=True)
class NodeView:
    """Read-only snapshot of a graph node.

    Instances are returned by ``DependencyGraph.get_node(...)`` and reflect the
    node's state at the time of the lookup. To observe subsequent mutations,
    re-fetch the view. Durable mutation is done via
    ``DependencyGraph.set_node_value(...)`` and ``set_node_metadata(...)``.
    """

    sheet: str
    column: str
    row: int
    formula: str | None
    normalized_formula: str | None
    value: Any
    is_leaf: bool
    metadata: Mapping[str, Any]

    @property
    def key(self) -> NodeKey:
        return _format_node_key(self.sheet, self.column, self.row)

    @property
    def address(self) -> str:
        return f"{self.column}{self.row}"

    @property
    def column_index(self) -> int:
        return int(fastpyxl.utils.cell.column_index_from_string(self.column))

    @property
    def value_type(self) -> ValueType:
        return _classify_value(self.value, is_leaf=self.is_leaf)


def node_to_view(node: Node) -> NodeView:
    """Build an immutable ``NodeView`` snapshot from a stored ``Node``."""
    return NodeView(
        sheet=node.sheet,
        column=node.column,
        row=node.row,
        formula=node.formula,
        normalized_formula=node.normalized_formula,
        value=node.value,
        is_leaf=node.is_leaf,
        metadata=MappingProxyType(dict(node.metadata)),
    )
