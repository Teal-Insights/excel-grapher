from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .node import NodeKey


@dataclass(frozen=True)
class GuardExpr:
    """Base type for conditional dependency guards."""


@dataclass(frozen=True)
class CellRef(GuardExpr):
    """A cell reference used in a condition."""

    key: NodeKey

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        return self.key


@dataclass(frozen=True)
class Literal(GuardExpr):
    """A literal value in a condition."""

    value: Any

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        v = self.value
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, str):
            return f'"{v}"'
        return str(v)


@dataclass(frozen=True)
class Compare(GuardExpr):
    """Comparison: left op right."""

    left: GuardExpr
    op: str
    right: GuardExpr

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        return f"{self.left}{self.op}{self.right}"


@dataclass(frozen=True)
class Not(GuardExpr):
    """Logical negation."""

    operand: GuardExpr

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        return f"NOT({self.operand})"


@dataclass(frozen=True)
class And(GuardExpr):
    """Logical AND."""

    operands: tuple[GuardExpr, ...]

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        inner = ",".join(str(o) for o in self.operands)
        return f"AND({inner})"


@dataclass(frozen=True)
class Or(GuardExpr):
    """Logical OR."""

    operands: tuple[GuardExpr, ...]

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        inner = ",".join(str(o) for o in self.operands)
        return f"OR({inner})"


def or_guard(a: GuardExpr, b: GuardExpr) -> GuardExpr:
    """
    Combine two guards with OR, flattening nested ORs.
    """
    ops: list[GuardExpr] = []
    if isinstance(a, Or):
        ops.extend(a.operands)
    else:
        ops.append(a)
    if isinstance(b, Or):
        ops.extend(b.operands)
    else:
        ops.append(b)
    return Or(tuple(ops))

