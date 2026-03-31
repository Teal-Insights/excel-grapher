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


@dataclass(frozen=True)
class GuardConstraints:
    """
    A minimal, conservative constraint set derived from a conjunction of guards.

    This is used to check whether a set of guard expressions is internally consistent
    (e.g., it can't contain both X=0 and X=1 at the same time).
    """

    equalities: tuple[tuple[NodeKey, Any], ...] = ()
    inequalities: tuple[tuple[NodeKey, tuple[Any, ...]], ...] = ()
    opaque: tuple[str, ...] = ()

    def add(self, g: GuardExpr) -> GuardConstraints | None:
        """
        Return a new GuardConstraints with g conjoined, or None if inconsistent.

        Only a small subset of GuardExpr forms participate in consistency checking:
        - Compare(CellRef(key), "=", Literal(v))
        - Compare(CellRef(key), "<>", Literal(v))
        - Not(Compare(...)) is rewritten when possible
        - And(...) is flattened into its operands
        Everything else is tracked as opaque (string form) without consistency checks.
        """

        def flatten(expr: GuardExpr) -> list[GuardExpr]:
            if isinstance(expr, And):
                out: list[GuardExpr] = []
                for o in expr.operands:
                    out.extend(flatten(o))
                return out
            return [expr]

        eq: dict[NodeKey, Any] = dict(self.equalities)
        ne: dict[NodeKey, set[Any]] = {k: set(vs) for k, vs in self.inequalities}
        opaque: set[str] = set(self.opaque)

        for expr in flatten(g):
            expr2: GuardExpr = expr
            if isinstance(expr2, Not) and isinstance(expr2.operand, Compare):
                c = expr2.operand
                if c.op == "=":
                    expr2 = Compare(left=c.left, op="<>", right=c.right)
                elif c.op == "<>":
                    expr2 = Compare(left=c.left, op="=", right=c.right)

            if (
                isinstance(expr2, Compare)
                and isinstance(expr2.left, CellRef)
                and isinstance(expr2.right, Literal)
            ):
                key = expr2.left.key
                val = expr2.right.value
                if expr2.op == "=":
                    existing = eq.get(key)
                    if existing is not None and existing != val:
                        return None
                    if key in ne and val in ne[key]:
                        return None
                    eq[key] = val
                    continue
                if expr2.op == "<>":
                    existing = eq.get(key)
                    if existing is not None and existing == val:
                        return None
                    ne.setdefault(key, set()).add(val)
                    continue

            opaque.add(str(expr2))

        eq_items = tuple(sorted(eq.items(), key=lambda kv: kv[0]))
        ne_items = tuple(
            sorted(((k, tuple(sorted(vs))) for k, vs in ne.items()), key=lambda kv: kv[0])
        )
        opaque_items = tuple(sorted(opaque))
        return GuardConstraints(equalities=eq_items, inequalities=ne_items, opaque=opaque_items)
