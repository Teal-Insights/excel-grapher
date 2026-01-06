from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum, auto
from typing import Any

import openpyxl.utils.cell

NodeKey = str  # Always in the form "SheetName!A1"


class ValueType(Enum):
    NUMBER = auto()
    STRING = auto()
    BOOLEAN = auto()
    ERROR = auto()
    DATETIME = auto()
    EMPTY = auto()
    UNKNOWN = auto()


@dataclass
class Node:
    sheet: str
    column: str
    row: int
    formula: str | None
    normalized_formula: str | None
    value: Any
    is_leaf: bool
    key_override: NodeKey | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def key(self) -> NodeKey:
        if self.key_override is not None:
            return self.key_override
        return f"{self.sheet}!{self.column}{self.row}"

    @property
    def address(self) -> str:
        return f"{self.column}{self.row}"

    @property
    def column_index(self) -> int:
        return int(openpyxl.utils.cell.column_index_from_string(self.column))

    @property
    def value_type(self) -> ValueType:
        v = self.value
        if v is None:
            return ValueType.EMPTY if self.is_leaf else ValueType.UNKNOWN
        # Must check bool before int because bool is a subclass of int in Python.
        if isinstance(v, bool):
            return ValueType.BOOLEAN
        if isinstance(v, (int, float)):
            return ValueType.NUMBER
        if isinstance(v, datetime):
            return ValueType.DATETIME
        if isinstance(v, str):
            if v.startswith("#"):
                return ValueType.ERROR
            return ValueType.STRING
        return ValueType.UNKNOWN

