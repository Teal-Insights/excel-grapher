from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from fastpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_grapher.core.address_keys import RangeKey, format_cell_key

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
class RangeRef(GuardExpr):
    """A multi-cell range used in an array-context condition.

    Array-context `IF` (`SUM(IF(A1:A10>0,B1:B10,0))` and friends) evaluates its
    condition element-wise, so a range in the condition is not a value: it is a
    placeholder for *the element aligned with the value element being guarded*.
    A `RangeRef` is therefore a **template** node — `instantiate_element_guard`
    resolves it to a `CellRef` per element, and only the resolved scalar guards
    are ever attached to graph edges.

    Attributes:
        key: Canonical range address, e.g. `Sheet1!A1:A10`.
    """

    key: NodeKey

    @property
    def shape(self) -> tuple[int, int]:
        """Return the range's `(n_rows, n_cols)` extent."""
        rk = RangeKey(self.key)
        n_rows = rk.max_row - rk.min_row + 1
        n_cols = column_index_from_string(rk.max_col) - column_index_from_string(rk.min_col) + 1
        return n_rows, n_cols

    def element(self, row_offset: int, col_offset: int) -> CellRef:
        """Return the `CellRef` at `(row_offset, col_offset)` within the range.

        Raises:
            IndexError: If the offsets fall outside the range.
        """
        n_rows, n_cols = self.shape
        if not (0 <= row_offset < n_rows and 0 <= col_offset < n_cols):
            raise IndexError(
                f"Element ({row_offset}, {col_offset}) is outside range {self.key} "
                f"of shape {n_rows}x{n_cols}"
            )
        rk = RangeKey(self.key)
        column = get_column_letter(column_index_from_string(rk.min_col) + col_offset)
        return CellRef(format_cell_key(rk.sheet, column, rk.min_row + row_offset))

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


def canonicalize_guard(expr: GuardExpr) -> GuardExpr:
    """Return a canonicalized guard expression for conservative symbolic reasoning.

    Canonicalization is intentionally minimal:
    - recurse through And / Or / Compare / Not
    - eliminate double negation: Not(Not(x)) -> x
    """
    if isinstance(expr, Compare):
        left = canonicalize_guard(expr.left)
        right = canonicalize_guard(expr.right)
        if left is expr.left and right is expr.right:
            return expr
        return Compare(left=left, op=expr.op, right=right)
    if isinstance(expr, And):
        ops = tuple(canonicalize_guard(o) for o in expr.operands)
        return And(ops)
    if isinstance(expr, Or):
        ops = tuple(canonicalize_guard(o) for o in expr.operands)
        return Or(ops)
    if isinstance(expr, Not):
        operand = canonicalize_guard(expr.operand)
        if isinstance(operand, Not):
            return canonicalize_guard(operand.operand)
        if operand is expr.operand:
            return expr
        return Not(operand=operand)
    return expr


def guard_range_shape(expr: GuardExpr) -> tuple[int, int] | None:
    """Return the common `(n_rows, n_cols)` of the guard's `RangeRef`s.

    Returns `None` when the guard holds no `RangeRef` (a plain scalar guard) or
    when its `RangeRef`s disagree on shape, since elements of differently shaped
    ranges cannot be aligned.
    """
    shapes = {r.shape for r in _collect_range_refs(expr)}
    if len(shapes) != 1:
        return None
    return next(iter(shapes))


def _collect_range_refs(expr: GuardExpr) -> list[RangeRef]:
    if isinstance(expr, RangeRef):
        return [expr]
    if isinstance(expr, Compare):
        return _collect_range_refs(expr.left) + _collect_range_refs(expr.right)
    if isinstance(expr, Not):
        return _collect_range_refs(expr.operand)
    if isinstance(expr, (And, Or)):
        out: list[RangeRef] = []
        for operand in expr.operands:
            out.extend(_collect_range_refs(operand))
        return out
    return []


def instantiate_element_guard(
    expr: GuardExpr, *, row_offset: int, col_offset: int
) -> GuardExpr | None:
    """Resolve an array-context guard template for one element.

    Every `RangeRef` is replaced by its element at `(row_offset, col_offset)`;
    scalar operands are left alone (they broadcast across the array).

    Returns:
        The scalar guard for that element, or `None` when the offsets fall
        outside one of the ranges (nothing sound can be said about the element).
    """
    if isinstance(expr, RangeRef):
        try:
            return expr.element(row_offset, col_offset)
        except IndexError:
            return None
    if isinstance(expr, Compare):
        left = instantiate_element_guard(expr.left, row_offset=row_offset, col_offset=col_offset)
        right = instantiate_element_guard(expr.right, row_offset=row_offset, col_offset=col_offset)
        if left is None or right is None:
            return None
        return Compare(left=left, op=expr.op, right=right)
    if isinstance(expr, Not):
        operand = instantiate_element_guard(
            expr.operand, row_offset=row_offset, col_offset=col_offset
        )
        return None if operand is None else Not(operand)
    if isinstance(expr, (And, Or)):
        operands: list[GuardExpr] = []
        for operand in expr.operands:
            resolved = instantiate_element_guard(
                operand, row_offset=row_offset, col_offset=col_offset
            )
            if resolved is None:
                return None
            operands.append(resolved)
        return And(tuple(operands)) if isinstance(expr, And) else Or(tuple(operands))
    return expr


def and_guard(a: GuardExpr, b: GuardExpr) -> GuardExpr:
    """Combine two guards with AND, flattening nested ANDs.

    `Literal(True)` is the AND identity and is dropped from the result.
    """
    ops: list[GuardExpr] = []
    for g in (a, b):
        if isinstance(g, And):
            ops.extend(g.operands)
        else:
            ops.append(g)
    ops = [g for g in ops if g != Literal(True)]
    if not ops:
        return Literal(True)
    if len(ops) == 1:
        return ops[0]
    return And(tuple(ops))


def or_guard(a: GuardExpr, b: GuardExpr) -> GuardExpr:
    """Combine two guards with OR, flattening nested ORs."""
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
    """A minimal, conservative constraint set derived from a conjunction of guards.

    This is used to check whether a set of guard expressions is internally consistent
    (e.g., it can't contain both X=0 and X=1 at the same time).
    """

    equalities: tuple[tuple[NodeKey, Any], ...] = ()
    inequalities: tuple[tuple[NodeKey, tuple[Any, ...]], ...] = ()
    opaque: tuple[str, ...] = ()

    def add(self, g: GuardExpr) -> GuardConstraints | None:
        """Return a new GuardConstraints with g conjoined, or None if inconsistent.

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

        for expr in flatten(canonicalize_guard(g)):
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
