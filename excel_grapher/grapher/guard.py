from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from fractions import Fraction
from heapq import heappop, heappush
from math import isfinite
from typing import Any
from weakref import WeakValueDictionary

from fastpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_grapher.core.address_keys import (
    CellKey,
    RangeKey,
    format_cell_key,
    format_range_key,
    parse_node_key,
)
from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    is_union_domain,
    normalize_cell_type_env_key,
)

from .node import NodeKey


@dataclass(frozen=True, slots=True, weakref_slot=True)
class GuardExpr:
    """Base type for conditional dependency guards."""


@dataclass(frozen=True, slots=True, weakref_slot=True)
class CellRef(GuardExpr):
    """A cell reference used in a condition."""

    key: NodeKey

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        return self.key


@dataclass(frozen=True, slots=True, weakref_slot=True)
class RangeRef(GuardExpr):
    """A multi-cell range used in an array-context condition.

    Array-context `IF` (`SUM(IF(A1:A10>0,B1:B10,0))` and friends) evaluates its
    condition element-wise, so a range in the condition is not a value: it is a
    placeholder for *the element aligned with the value element being guarded*.
    A `RangeRef` is therefore a **template** node — `instantiate_element_guard`
    resolves it to a `CellRef` per element, and only the resolved scalar guards
    are ever attached to graph edges.

    Those scalar guards mean the aligned value cell can change the result, not
    that Excel omits it from the calc chain. The engine still treats the whole
    value range as a precedent; the graph drops the edge from a value walk
    when the aligned condition is false so evaluation stays Excel-correct
    with fewer edges.

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
        cell = intern_guard(CellRef(format_cell_key(rk.sheet, column, rk.min_row + row_offset)))
        assert isinstance(cell, CellRef)
        return cell

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        return self.key


@dataclass(frozen=True, slots=True, weakref_slot=True)
class Literal(GuardExpr):
    """A literal value in a condition."""

    value: Any

    def __eq__(self, other: object) -> bool:
        # Bool is a subclass of int; `True == 1` must not collapse distinct literals.
        if not isinstance(other, Literal):
            return NotImplemented
        return type(self.value) is type(other.value) and self.value == other.value

    def __hash__(self) -> int:
        return hash((type(self.value), self.value))

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        v = self.value
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, str):
            return f'"{v}"'
        return str(v)


@dataclass(frozen=True, slots=True, weakref_slot=True)
class Arith(GuardExpr):
    """Numeric `left + right` or `left - right` operand of a comparison.

    Attributes:
        left: Left operand (`CellRef`, numeric `Literal`, `Arith` or `Neg`).
        op: `"+"` or `"-"`.
        right: Right operand.
    """

    left: GuardExpr
    op: str
    right: GuardExpr

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        right = f"({self.right})" if isinstance(self.right, Arith) else str(self.right)
        return f"{self.left}{self.op}{right}"


@dataclass(frozen=True, slots=True, weakref_slot=True)
class Neg(GuardExpr):
    """Numeric unary minus of a comparison operand."""

    operand: GuardExpr

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        if isinstance(self.operand, (Arith, Neg)):
            return f"-({self.operand})"
        return f"-{self.operand}"


@dataclass(frozen=True, slots=True, weakref_slot=True)
class Compare(GuardExpr):
    """Comparison: left op right."""

    left: GuardExpr
    op: str
    right: GuardExpr

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        return f"{self.left}{self.op}{self.right}"


@dataclass(frozen=True, slots=True, weakref_slot=True)
class Not(GuardExpr):
    """Logical negation."""

    operand: GuardExpr

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        return f"NOT({self.operand})"


@dataclass(frozen=True, slots=True, weakref_slot=True)
class And(GuardExpr):
    """Logical AND."""

    operands: tuple[GuardExpr, ...]

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        inner = ",".join(str(o) for o in self.operands)
        return f"AND({inner})"


@dataclass(frozen=True, slots=True, weakref_slot=True)
class Or(GuardExpr):
    """Logical OR."""

    operands: tuple[GuardExpr, ...]

    def __str__(self) -> str:  # pragma: no cover (covered indirectly via exports)
        inner = ",".join(str(o) for o in self.operands)
        return f"OR({inner})"


# Weak hash-cons table: structural keys so the key does not keep the value alive.
# Entries disappear once no graph (or other strong ref) retains the expression.
_GUARD_POOL: WeakValueDictionary[tuple[Any, ...], GuardExpr] = WeakValueDictionary()


def _guard_intern_key(expr: GuardExpr) -> tuple[Any, ...]:
    """Return a structural pool key that does not strongly reference `expr`.

    Composite nodes key children by `id()`; callers must intern children first so
    equal subtrees share identity and thus the same child ids.
    """
    if isinstance(expr, CellRef):
        return (CellRef, expr.key)
    if isinstance(expr, RangeRef):
        return (RangeRef, expr.key)
    if isinstance(expr, Literal):
        return (Literal, type(expr.value), expr.value)
    if isinstance(expr, Compare):
        return (Compare, id(expr.left), expr.op, id(expr.right))
    if isinstance(expr, Arith):
        return (Arith, id(expr.left), expr.op, id(expr.right))
    if isinstance(expr, Neg):
        return (Neg, id(expr.operand))
    if isinstance(expr, Not):
        return (Not, id(expr.operand))
    if isinstance(expr, And):
        return (And, *(id(o) for o in expr.operands))
    if isinstance(expr, Or):
        return (Or, *(id(o) for o in expr.operands))
    return (type(expr), id(expr))


def intern_guard(expr: GuardExpr) -> GuardExpr:
    """Return the canonical shared instance for `expr` (hash-cons / intern).

    Child nodes are interned first so equal subtrees are shared by identity as
    well as equality. Distinct-but-equal trees constructed independently (for
    example the same `IF` condition on two cells) collapse to one object.

    The intern pool holds values weakly: expressions retained only by the pool
    are eligible for collection once no graph (or other strong ref) keeps them.
    """
    if isinstance(expr, Compare):
        left = intern_guard(expr.left)
        right = intern_guard(expr.right)
        if left is not expr.left or right is not expr.right:
            expr = Compare(left=left, op=expr.op, right=right)
    elif isinstance(expr, Arith):
        left = intern_guard(expr.left)
        right = intern_guard(expr.right)
        if left is not expr.left or right is not expr.right:
            expr = Arith(left=left, op=expr.op, right=right)
    elif isinstance(expr, Neg):
        operand = intern_guard(expr.operand)
        if operand is not expr.operand:
            expr = Neg(operand=operand)
    elif isinstance(expr, Not):
        operand = intern_guard(expr.operand)
        if operand is not expr.operand:
            expr = Not(operand=operand)
    elif isinstance(expr, And):
        ops = tuple(intern_guard(o) for o in expr.operands)
        if any(a is not b for a, b in zip(ops, expr.operands, strict=True)):
            expr = And(ops)
    elif isinstance(expr, Or):
        ops = tuple(intern_guard(o) for o in expr.operands)
        if any(a is not b for a, b in zip(ops, expr.operands, strict=True)):
            expr = Or(ops)

    key = _guard_intern_key(expr)
    cached = _GUARD_POOL.get(key)
    if cached is not None:
        return cached
    _GUARD_POOL[key] = expr
    return expr


def guard_intern_pool_size() -> int:
    """Return the number of live entries in the guard intern pool."""
    return len(_GUARD_POOL)


def clear_guard_intern_pool() -> None:
    """Drop all intern-pool entries and re-seed `Literal(True)`.

    Expressions still held by graphs remain alive but will not be shared with
    newly interned equals until those equals are interned again. Intended for
    batch jobs that extract many independent workbooks in one process.
    """
    global _LITERAL_TRUE
    _GUARD_POOL.clear()
    _LITERAL_TRUE = intern_guard(Literal(True))


_LITERAL_TRUE = intern_guard(Literal(True))


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
            return intern_guard(expr)
        return intern_guard(Compare(left=left, op=expr.op, right=right))
    if isinstance(expr, And):
        ops = tuple(canonicalize_guard(o) for o in expr.operands)
        return intern_guard(And(ops))
    if isinstance(expr, Or):
        ops = tuple(canonicalize_guard(o) for o in expr.operands)
        return intern_guard(Or(ops))
    if isinstance(expr, Not):
        operand = canonicalize_guard(expr.operand)
        if isinstance(operand, Not):
            return canonicalize_guard(operand.operand)
        if operand is expr.operand:
            return intern_guard(expr)
        return intern_guard(Not(operand=operand))
    return intern_guard(expr)


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
    if isinstance(expr, (Compare, Arith)):
        return _collect_range_refs(expr.left) + _collect_range_refs(expr.right)
    if isinstance(expr, (Not, Neg)):
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
    if isinstance(expr, (Compare, Arith)):
        left = instantiate_element_guard(expr.left, row_offset=row_offset, col_offset=col_offset)
        right = instantiate_element_guard(expr.right, row_offset=row_offset, col_offset=col_offset)
        if left is None or right is None:
            return None
        return intern_guard(type(expr)(left=left, op=expr.op, right=right))
    if isinstance(expr, (Not, Neg)):
        operand = instantiate_element_guard(
            expr.operand, row_offset=row_offset, col_offset=col_offset
        )
        return None if operand is None else intern_guard(type(expr)(operand))
    if isinstance(expr, (And, Or)):
        operands: list[GuardExpr] = []
        for operand in expr.operands:
            resolved = instantiate_element_guard(
                operand, row_offset=row_offset, col_offset=col_offset
            )
            if resolved is None:
                return None
            operands.append(resolved)
        return intern_guard(And(tuple(operands)) if isinstance(expr, And) else Or(tuple(operands)))
    return intern_guard(expr)


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
    ops = [g for g in ops if g != _LITERAL_TRUE]
    if not ops:
        return _LITERAL_TRUE
    if len(ops) == 1:
        return intern_guard(ops[0])
    return intern_guard(And(tuple(ops)))


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
    return intern_guard(Or(tuple(ops)))


def _rewrite_range_ref_key(key: NodeKey, old_key: NodeKey, new_key: NodeKey) -> NodeKey:
    """Rewrite bounding-box corners of a range key; interior occupancy is ignored."""
    try:
        rk = key if isinstance(key, RangeKey) else RangeKey(str(key))
        start = format_cell_key(rk.sheet, rk.min_col, rk.min_row)
        end = format_cell_key(rk.sheet, rk.max_col, rk.max_row)
    except ValueError:
        return key
    if start != old_key and end != old_key:
        return key
    dest = parse_node_key(str(new_key))
    if not isinstance(dest, CellKey):
        raise ValueError(
            f"range guard endpoint rewrite requires a cell destination, got {new_key!r}"
        )
    start_sheet, start_col, start_row = rk.sheet, rk.min_col, rk.min_row
    end_sheet, end_col, end_row = rk.sheet, rk.max_col, rk.max_row
    if start == old_key:
        start_sheet, start_col, start_row = dest.sheet, dest.column, dest.row
    if end == old_key:
        end_sheet, end_col, end_row = dest.sheet, dest.column, dest.row
    if start_sheet != end_sheet:
        raise ValueError(
            f"rewriting range guard {key} would split sheets ({start_sheet!r} vs {end_sheet!r})"
        )
    return format_range_key(start_sheet, f"{start_col}{start_row}", f"{end_col}{end_row}")


def rewrite_guard_keys(expr: GuardExpr, old_key: NodeKey, new_key: NodeKey) -> GuardExpr:
    """Replace `old_key` with `new_key` in cell refs and range endpoints.

    Interior occupancy of a `RangeRef` is not a rewrite site: only bounding-box
    corners that equal `old_key` change. Composite nodes are rebuilt and interned.

    Args:
        expr: Guard tree to rewrite.
        old_key: Cell key to replace.
        new_key: Replacement cell key.

    Returns:
        Interned guard; `expr` when no cell or range-endpoint key changed.

    Raises:
        ValueError: If rewriting a `RangeRef` would split it across sheets.
    """
    if old_key == new_key:
        return intern_guard(expr)

    def walk(node: GuardExpr) -> GuardExpr:
        if isinstance(node, CellRef):
            if node.key != old_key:
                return intern_guard(node)
            return intern_guard(CellRef(new_key))
        if isinstance(node, RangeRef):
            rewritten = _rewrite_range_ref_key(node.key, old_key, new_key)
            if rewritten == node.key:
                return intern_guard(node)
            return intern_guard(RangeRef(rewritten))
        if isinstance(node, (Compare, Arith)):
            left = walk(node.left)
            right = walk(node.right)
            if left is node.left and right is node.right:
                return intern_guard(node)
            return intern_guard(type(node)(left=left, op=node.op, right=right))
        if isinstance(node, (Not, Neg)):
            operand = walk(node.operand)
            if operand is node.operand:
                return intern_guard(node)
            return intern_guard(type(node)(operand=operand))
        if isinstance(node, And):
            ops = tuple(walk(o) for o in node.operands)
            if all(a is b for a, b in zip(ops, node.operands, strict=True)):
                return intern_guard(node)
            return intern_guard(And(ops))
        if isinstance(node, Or):
            ops = tuple(walk(o) for o in node.operands)
            if all(a is b for a, b in zip(ops, node.operands, strict=True)):
                return intern_guard(node)
            return intern_guard(Or(ops))
        return intern_guard(node)

    return walk(expr)


def _constraint_key(key: NodeKey) -> NodeKey:
    """Canonicalize a guard cell key so quoted graph keys match `CellTypeEnv` keys."""
    try:
        return normalize_cell_type_env_key(str(key))
    except (ValueError, IndexError):
        return key


def _find_parent(parent: dict[NodeKey, NodeKey], key: NodeKey) -> NodeKey:
    root = key
    seen: list[NodeKey] = []
    while parent.get(root, root) != root:
        seen.append(root)
        root = parent[root]
    for node in seen:
        parent[node] = root
    return root


def _value_allowed_by_env(env: CellTypeEnv | None, key: NodeKey, val: Any) -> bool:
    """Return whether `val` is allowed for `key` under `env` (unknown -> True)."""
    if env is None:
        return True
    cell_type = env.get(_constraint_key(key))
    if cell_type is None:
        return True
    if is_union_domain(cell_type):
        if _enum_member(cell_type, val):
            return True
        return _interval_allows(cell_type, val)
    if cell_type.enum is not None and val not in cell_type.enum.values:
        return False
    if isinstance(val, bool) or not isinstance(val, (int, float)):
        return True
    if cell_type.interval is not None:
        lo, hi = cell_type.interval.min, cell_type.interval.max
        if lo is not None and val < lo:
            return False
        if hi is not None and val > hi:
            return False
    if cell_type.real_interval is not None:
        lo_r, hi_r = cell_type.real_interval.min, cell_type.real_interval.max
        if lo_r is not None and float(val) < lo_r:
            return False
        if hi_r is not None and float(val) > hi_r:
            return False
    return True


def _enum_remaining(env: CellTypeEnv | None, key: NodeKey, forbidden: set[Any]) -> bool:
    """Return False when every enum member for `key` is forbidden."""
    if env is None:
        return True
    cell_type = env.get(_constraint_key(key))
    if cell_type is None or cell_type.enum is None:
        return True
    if is_union_domain(cell_type):
        if cell_type.enum.values - forbidden:
            return True
        return _union_interval_remains(cell_type, forbidden)
    return bool(cell_type.enum.values - forbidden)


def _enum_member(cell_type: CellType, value: Any) -> bool:
    """Return whether `value` matches an enum member without bool/int confusion."""
    if cell_type.enum is None:
        return False
    return any(type(value) is type(item) and value == item for item in cell_type.enum.values)


def _interval_allows(cell_type: CellType, value: Any) -> bool:
    """Return whether `value` lies in the cell's integer or real interval."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    if cell_type.interval is not None:
        low, high = cell_type.interval.min, cell_type.interval.max
        if low is not None and value < low:
            return False
        if high is not None and value > high:
            return False
    if cell_type.real_interval is not None:
        low, high = cell_type.real_interval.min, cell_type.real_interval.max
        if low is not None and float(value) < low:
            return False
        if high is not None and float(value) > high:
            return False
    return cell_type.interval is not None or cell_type.real_interval is not None


def _union_interval_remains(cell_type: CellType, forbidden: set[Any]) -> bool:
    """Return whether a union's interval still contains a value outside `forbidden`."""
    interval = cell_type.interval
    if interval is not None and interval.min is not None and interval.max is not None:
        span = interval.max - interval.min + 1
        if span > len(forbidden):
            return True
        return any(value not in forbidden for value in range(interval.min, interval.max + 1))
    if interval is not None:
        return True
    real = cell_type.real_interval
    if real is None:
        return False
    if real.min is not None and real.max is not None and real.min == real.max:
        return real.min not in forbidden
    return True


# Difference bounds `x - y <= c + e*eps`, keyed `(x, y)`, where `e` is `-1` for a
# strict bound and sums of weights add both parts, so tuple order is bound order.
# The empty key stands for the constant 0, so `x <= c` is `(x, _ZERO)`.
_ZERO: NodeKey = ""
_Num = int | Fraction
_Weight = tuple[_Num, int]
_NO_WEIGHT: _Weight = (0, 0)
# A linear row `sum(coef * cell) + const` (`< 0` when strict, else `<= 0`).
_LinearRow = tuple[dict[NodeKey, int], _Num, bool]

_ORDERED_COMPLEMENT = {"<": ">=", "<=": ">", ">": "<=", ">=": "<"}
# Rows with more cells than this are left opaque rather than projected.
_MAX_ROW_CELLS = 8


def _exact(v: float) -> _Num:
    """Return `v` as an `int` when integral, else as the `Fraction` of its decimal text.

    Decimal text keeps `0.1 + 0.2 == 0.3`, matching the formula author's intent.
    """
    if isinstance(v, int):
        return v
    f = Fraction(repr(v))
    return f.numerator if f.denominator == 1 else f


def _numeric_domain(
    env: CellTypeEnv | None, key: NodeKey
) -> tuple[_Num | None, _Num | None] | None:
    """Return `(lo, hi)` numeric bounds for `key`, or `None` if it may hold non-numbers.

    Cells without a declared type are assumed numeric (blanks compare as 0).
    Text and booleans sort above every number in Excel comparisons, so a cell
    whose domain admits them cannot be ordered as a real number.
    """
    cell_type = None if env is None else env.get(key)
    if cell_type is None:
        return None, None
    if cell_type.kind in (CellKind.STRING, CellKind.BOOL, CellKind.ERROR):
        return None
    points: list[_Num] = []
    if cell_type.enum is not None:
        for v in cell_type.enum.values:
            if isinstance(v, bool) or not isinstance(v, (int, float)) or not isfinite(v):
                return None
            points.append(_exact(v))
    los: list[_Num | None] = []
    his: list[_Num | None] = []
    for dom in (cell_type.interval, cell_type.real_interval):
        if dom is not None:
            los.append(None if dom.min is None else _exact(dom.min))
            his.append(None if dom.max is None else _exact(dom.max))
    known_lo = [v for v in los if v is not None]
    known_hi = [v for v in his if v is not None]
    if is_union_domain(cell_type):
        # Enum or interval: the hull of both.
        lo = None if None in los else min(points + known_lo)
        hi = None if None in his else max(points + known_hi)
        return lo, hi
    if points:
        return min(points), max(points)
    return (max(known_lo) if known_lo else None, min(known_hi) if known_hi else None)


def _linear_terms(
    expr: GuardExpr, env: CellTypeEnv | None
) -> tuple[dict[NodeKey, int], _Num] | None:
    """Return `expr` as `(cell coefficients, constant)`, or `None` if not linear-numeric."""
    if isinstance(expr, CellRef):
        key = _constraint_key(expr.key)
        return None if _numeric_domain(env, key) is None else ({key: 1}, 0)
    if isinstance(expr, Literal):
        v = expr.value
        if isinstance(v, bool) or not isinstance(v, (int, float)) or not isfinite(v):
            return None
        return {}, _exact(v)
    if isinstance(expr, Neg):
        inner = _linear_terms(expr.operand, env)
        if inner is None:
            return None
        return {k: -a for k, a in inner[0].items()}, -inner[1]
    if isinstance(expr, Arith) and expr.op in ("+", "-"):
        left = _linear_terms(expr.left, env)
        right = _linear_terms(expr.right, env)
        if left is None or right is None:
            return None
        sign = 1 if expr.op == "+" else -1
        coefs = dict(left[0])
        for k, a in right[0].items():
            coefs[k] = coefs.get(k, 0) + sign * a
        return {k: a for k, a in coefs.items() if a}, left[1] + sign * right[1]
    return None


def _ordered_rows(expr: Compare, env: CellTypeEnv | None) -> list[_LinearRow] | None:
    """Return the linear rows of an ordered or `=` comparison, or `None`."""
    left = _linear_terms(expr.left, env)
    right = _linear_terms(expr.right, env)
    if left is None or right is None:
        return None
    coefs = dict(left[0])
    for k, a in right[0].items():
        coefs[k] = coefs.get(k, 0) - a
    diff = ({k: a for k, a in coefs.items() if a}, left[1] - right[1])
    neg = ({k: -a for k, a in diff[0].items()}, -diff[1])
    rows = {
        "<": [(*diff, True)],
        "<=": [(*diff, False)],
        ">": [(*neg, True)],
        ">=": [(*neg, False)],
        "=": [(*diff, False), (*neg, False)],
    }.get(expr.op)
    return rows


def _project_row(
    row: _LinearRow, env: CellTypeEnv | None
) -> tuple[list[tuple[NodeKey, NodeKey, _Weight]], bool, bool]:
    """Project a linear row onto difference bounds.

    Every pair of a `+1` cell and a `-1` cell (either may be absent) is kept and
    the remaining cells are replaced by their domain bound in the weakening
    direction, so each emitted bound is implied by the row.

    Returns:
        `(bounds, exact, violated)`: bounds as `(x, y, weight)` for
        `x - y <= weight`; `exact` when some projection kept every cell; `violated`
        when a constant-only projection is already false.
    """
    coefs, const, strict = row
    if len(coefs) > _MAX_ROW_CELLS:
        return [], False, False
    domains = {k: _numeric_domain(env, k) or (None, None) for k in coefs}
    plus = [k for k, a in coefs.items() if a == 1]
    minus = [k for k, a in coefs.items() if a == -1]
    out: list[tuple[NodeKey, NodeKey, _Weight]] = []
    exact = False
    for p in [*plus, None]:
        for n in [*minus, None]:
            slack = const
            for k, a in coefs.items():
                if k in (p, n):
                    continue
                lo, hi = domains[k]
                bound = lo if a > 0 else hi
                if bound is None:
                    break
                slack += a * bound
            else:
                exact = exact or len(coefs) == (p is not None) + (n is not None)
                if p is None and n is None:
                    if slack > 0 or (slack == 0 and strict):
                        return [], exact, True
                    continue
                out.append((p or _ZERO, n or _ZERO, (-slack, -1 if strict else 0)))
    return out, exact, False


def _add_difference_bound(
    bounds: dict[tuple[NodeKey, NodeKey], _Weight],
    potential: dict[NodeKey, _Weight],
    x: NodeKey,
    y: NodeKey,
    w: _Weight,
) -> bool:
    """Conjoin `x - y <= w`; return False when the bounds become infeasible.

    `potential` is kept a satisfying assignment (`p[x] - p[y] <= w` for every
    bound). A bound it already satisfies costs O(1); otherwise the violated
    potentials are repaired Dijkstra-style over the affected nodes only, and
    having to lower `y` itself means a negative cycle (Cotton & Maler, 2006).
    """
    if x == y:
        return w >= _NO_WEIGHT
    old = bounds.get((x, y))
    if old is not None and old <= w:
        return True
    bounds[(x, y)] = w
    # A fresh cell has no other bounds, so it can take any satisfying value.
    if x not in potential:
        py = potential.setdefault(y, _NO_WEIGHT)
        potential[x] = (py[0] + w[0], py[1] + w[1])
        return True
    if y not in potential:
        px = potential[x]
        potential[y] = (px[0] - w[0], px[1] - w[1])
        return True
    px, py = potential[x], potential[y]
    first = (py[0] + w[0] - px[0], py[1] + w[1] - px[1])
    if first >= _NO_WEIGHT:
        return True
    succ: dict[NodeKey, list[tuple[NodeKey, _Weight]]] = {}
    for (a, b), wt in bounds.items():
        succ.setdefault(b, []).append((a, wt))
    lowered: dict[NodeKey, _Weight] = {}
    best: dict[NodeKey, _Weight] = {x: first}
    heap: list[tuple[_Weight, NodeKey]] = [(first, x)]
    while heap:
        delta, s = heappop(heap)
        if s in lowered or best.get(s) != delta:
            continue
        if s == y:
            return False
        ps = potential[s]
        new_s = (ps[0] + delta[0], ps[1] + delta[1])
        lowered[s] = new_s
        for t, wt in succ.get(s, ()):
            if t in lowered:
                continue
            pt = potential[t]
            cand = (new_s[0] + wt[0] - pt[0], new_s[1] + wt[1] - pt[1])
            if cand < _NO_WEIGHT and (t not in best or cand < best[t]):
                best[t] = cand
                heappush(heap, (cand, t))
    potential.update(lowered)
    return True


@dataclass(frozen=True)
class GuardConstraints:
    """A minimal, conservative constraint set derived from a conjunction of guards.

    This is used to check whether a set of guard expressions is internally consistent
    (e.g., it can't contain both X=0 and X=1 at the same time).
    """

    equalities: tuple[tuple[NodeKey, Any], ...] = ()
    inequalities: tuple[tuple[NodeKey, tuple[Any, ...]], ...] = ()
    opaque: tuple[str, ...] = ()
    cell_parents: tuple[tuple[NodeKey, NodeKey], ...] = ()
    cell_ne: tuple[tuple[NodeKey, NodeKey], ...] = ()
    bounds: tuple[tuple[NodeKey, NodeKey, _Weight], ...] = ()
    # A satisfying assignment for `bounds`; path-dependent, so not part of identity.
    potential: tuple[tuple[NodeKey, _Weight], ...] = field(default=(), compare=False)

    def _frozen_bounds(
        self,
        bounds: Mapping[tuple[NodeKey, NodeKey], _Weight],
        potential: Mapping[NodeKey, _Weight],
    ) -> tuple[tuple[tuple[NodeKey, NodeKey, _Weight], ...], tuple[tuple[NodeKey, _Weight], ...]]:
        """Freeze bounds and potentials as sorted tuples.

        Unchanged entries reuse this state's tuples, so the many DFS states along
        a path share them and each state costs about a pointer per entry.
        """
        kept_bounds = {(t[0], t[1]): t for t in self.bounds}
        triples = []
        for key, w in bounds.items():
            t = kept_bounds.get(key)
            triples.append(t if t is not None and t[2] == w else (key[0], key[1], w))
        kept_pot = {pair[0]: pair for pair in self.potential}
        pairs = []
        for k, w in potential.items():
            pair = kept_pot.get(k)
            pairs.append(pair if pair is not None and pair[1] == w else (k, w))
        # `(x, y)` / keys are unique, so sorting never compares weights.
        return tuple(sorted(triples)), tuple(sorted(pairs))

    def seed_cell_type_env(self, env: CellTypeEnv) -> GuardConstraints | None:
        """Conjoin singleton enum domains as equalities.

        Interval and multi-value enum domains are not seeded; they only reject
        assignments during `add`.
        """
        out: GuardConstraints | None = self
        for key, cell_type in env.items():
            if cell_type.enum is None or len(cell_type.enum.values) != 1:
                continue
            if is_union_domain(cell_type):
                continue
            value = next(iter(cell_type.enum.values))
            assert out is not None
            out = out.add(
                Compare(left=CellRef(key=key), op="=", right=Literal(value=value)),
                cell_type_env=env,
            )
            if out is None:
                return None
        return out

    def add(
        self, g: GuardExpr, *, cell_type_env: CellTypeEnv | None = None
    ) -> GuardConstraints | None:
        """Return a new GuardConstraints with g conjoined, or None if inconsistent.

        Forms that participate in consistency checking:
        - Compare(CellRef(key), "=", Literal(v)) and the swapped operand order
        - Compare(CellRef(key), "<>", Literal(v)) and the swapped operand order
        - Compare(CellRef(a), "=", CellRef(b)) / "<>" (unification)
        - `<`, `<=`, `>`, `>=` and `=` between `+`/`-` sums of cells and numeric
          literals, as difference bounds `x - y <= c` checked for a negative
          cycle. Terms with more cells keep one `+1` and one `-1` cell and
          replace the rest by their `cell_type_env` interval bound.
        - Not(Compare(...)) is rewritten when possible
        - And(...) is flattened into its operands
        Everything else is tracked as opaque (string form) without consistency checks.

        When `cell_type_env` is provided, equalities outside a cell's enum or
        interval are inconsistent, and complementary inequalities that exhaust
        a finite enum are inconsistent.

        Ordered reasoning treats cells as real numbers unless `cell_type_env`
        admits text, booleans or errors for them (those sort above numbers).
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
        parent: dict[NodeKey, NodeKey] = dict(self.cell_parents)
        ne_pairs: set[tuple[NodeKey, NodeKey]] = set(self.cell_ne)
        # Difference bounds are over union-find roots and copied only when touched,
        # so guards without ordered comparisons pay nothing for them.
        bounds: dict[tuple[NodeKey, NodeKey], _Weight] | None = None
        potential: dict[NodeKey, _Weight] | None = None

        def open_bounds() -> tuple[dict[tuple[NodeKey, NodeKey], _Weight], dict[NodeKey, _Weight]]:
            nonlocal bounds, potential
            if bounds is None or potential is None:
                bounds = {(t[0], t[1]): t[2] for t in self.bounds}
                potential = dict(self.potential)
            return bounds, potential

        def in_bounds(root: NodeKey) -> bool:
            if potential is None:
                return any(k == root for k, _ in self.potential)
            return root in potential

        def numeric_eq_bounds(root: NodeKey) -> list[tuple[NodeKey, NodeKey, _Weight]]:
            val = eq.get(root)
            if isinstance(val, bool) or not isinstance(val, (int, float)) or not isfinite(val):
                return []
            v = _exact(val)
            return [(root, _ZERO, (v, 0)), (_ZERO, root, (-v, 0))]

        def add_bound(x: NodeKey, y: NodeKey, w: _Weight) -> bool:
            """Conjoin `x - y <= w` over cell keys (or `_ZERO`), mapped to roots."""
            b, pot = open_bounds()
            pending = [(x, y, w)]
            for k in (x, y):
                if k == _ZERO:
                    continue
                lo, hi = _numeric_domain(cell_type_env, k) or (None, None)
                if hi is not None:
                    pending.append((k, _ZERO, (hi, 0)))
                if lo is not None:
                    pending.append((_ZERO, k, (-lo, 0)))
            for px, py, pw in pending:
                rx = px if px == _ZERO else find(px)
                ry = py if py == _ZERO else find(py)
                for r in (rx, ry):
                    if r != _ZERO and r not in pot:
                        pending.extend(numeric_eq_bounds(r))
                if not _add_difference_bound(b, pot, rx, ry, pw):
                    return False
            return True

        def remap_bounds() -> bool:
            """Re-add every bound under current roots after a union."""
            b, pot = open_bounds()
            old = list(b.items())
            b.clear()
            pot.clear()
            for (x, y), w in old:
                rx = x if x == _ZERO else find(x)
                ry = y if y == _ZERO else find(y)
                if not _add_difference_bound(b, pot, rx, ry, w):
                    return False
            return all(
                _add_difference_bound(b, pot, *bound)
                for r in list(pot)
                if r != _ZERO
                for bound in numeric_eq_bounds(r)
            )

        def find(key: NodeKey) -> NodeKey:
            return _find_parent(parent, _constraint_key(key))

        def add_eq_lit(raw_key: NodeKey, val: Any) -> bool:
            key = find(raw_key)
            existing = eq.get(key)
            if existing is not None and existing != val:
                return False
            if key in ne and val in ne[key]:
                return False
            if not _value_allowed_by_env(cell_type_env, key, val):
                return False
            eq[key] = val
            if in_bounds(key):
                b, pot = open_bounds()
                return all(_add_difference_bound(b, pot, *bd) for bd in numeric_eq_bounds(key))
            return True

        def add_ne_lit(raw_key: NodeKey, val: Any) -> bool:
            key = find(raw_key)
            existing = eq.get(key)
            if existing is not None and existing == val:
                return False
            ne.setdefault(key, set()).add(val)
            return _enum_remaining(cell_type_env, key, ne[key])

        def add_eq_cells(a: NodeKey, b: NodeKey) -> bool:
            ra, rb = find(a), find(b)
            pair = (min(ra, rb), max(ra, rb))
            if ra == rb:
                return True
            if pair in ne_pairs:
                return False
            va, vb = eq.get(ra), eq.get(rb)
            if va is not None and vb is not None and va != vb:
                return False
            parent[rb] = ra
            if vb is not None:
                if ra in eq and eq[ra] != vb:
                    return False
                eq[ra] = vb
            if ra in ne and rb in ne:
                ne[ra] = ne[ra] | ne[rb]
            elif rb in ne:
                ne.setdefault(ra, set()).update(ne[rb])
            remapped: set[tuple[NodeKey, NodeKey]] = set()
            for x, y in ne_pairs:
                xx, yy = find(x), find(y)
                if xx == yy:
                    return False
                remapped.add((min(xx, yy), max(xx, yy)))
            ne_pairs.clear()
            ne_pairs.update(remapped)
            if (in_bounds(ra) or in_bounds(rb)) and not remap_bounds():
                return False
            return ra not in ne or _enum_remaining(cell_type_env, ra, ne[ra])

        def add_ordered(c: Compare) -> bool | None:
            """Conjoin `c` as difference bounds; `None` when `c` is not fully captured."""
            rows = _ordered_rows(c, cell_type_env)
            if rows is None:
                return None
            exact = True
            for row in rows:
                projected, row_exact, violated = _project_row(row, cell_type_env)
                if violated:
                    return False
                exact = exact and row_exact
                for x, y, w in projected:
                    if not add_bound(x, y, w):
                        return False
            return True if exact else None

        def add_ne_cells(a: NodeKey, b: NodeKey) -> bool:
            ra, rb = find(a), find(b)
            if ra == rb:
                return False
            va, vb = eq.get(ra), eq.get(rb)
            if va is not None and vb is not None and va == vb:
                return False
            ne_pairs.add((min(ra, rb), max(ra, rb)))
            return True

        for expr in flatten(canonicalize_guard(g)):
            expr2: GuardExpr = expr
            if isinstance(expr2, Not) and isinstance(expr2.operand, Compare):
                c = expr2.operand
                if c.op == "=":
                    expr2 = Compare(left=c.left, op="<>", right=c.right)
                elif c.op == "<>":
                    expr2 = Compare(left=c.left, op="=", right=c.right)
                elif c.op in _ORDERED_COMPLEMENT:
                    expr2 = Compare(left=c.left, op=_ORDERED_COMPLEMENT[c.op], right=c.right)

            captured = False
            if isinstance(expr2, Compare) and expr2.op in ("=", "<>"):
                left, right = expr2.left, expr2.right
                cell: CellRef | None = None
                lit: Literal | None = None
                other: CellRef | None = None
                if isinstance(left, CellRef) and isinstance(right, Literal):
                    cell, lit = left, right
                elif isinstance(right, CellRef) and isinstance(left, Literal):
                    cell, lit = right, left
                elif isinstance(left, CellRef) and isinstance(right, CellRef):
                    cell, other = left, right

                if cell is not None and lit is not None:
                    ok = (
                        add_eq_lit(cell.key, lit.value)
                        if expr2.op == "="
                        else add_ne_lit(cell.key, lit.value)
                    )
                    if not ok:
                        return None
                    captured = True
                if cell is not None and other is not None:
                    ok = (
                        add_eq_cells(cell.key, other.key)
                        if expr2.op == "="
                        else add_ne_cells(cell.key, other.key)
                    )
                    if not ok:
                        return None
                    captured = True

            # Captured `=` forms join the bounds lazily, only once their cell is ordered.
            if isinstance(expr2, Compare) and expr2.op != "<>" and not captured:
                ordered = add_ordered(expr2)
                if ordered is False:
                    return None
                captured = captured or ordered is True

            if not captured:
                opaque.add(str(expr2))

        eq_items = tuple(sorted(eq.items(), key=lambda kv: kv[0]))
        ne_items = tuple(
            sorted(((k, tuple(sorted(vs))) for k, vs in ne.items()), key=lambda kv: kv[0])
        )
        parent_items = tuple(sorted((k, p) for k, p in parent.items() if k != p))
        ne_pair_items = tuple(sorted(ne_pairs))
        bound_items, potential_items = (
            (self.bounds, self.potential)
            if bounds is None or potential is None
            else self._frozen_bounds(bounds, potential)
        )
        return GuardConstraints(
            equalities=eq_items,
            inequalities=ne_items,
            opaque=tuple(sorted(opaque)),
            cell_parents=parent_items,
            cell_ne=ne_pair_items,
            bounds=bound_items,
            potential=potential_items,
        )


def rewrite_guard_aliases(expr: GuardExpr, aliases: Mapping[NodeKey, NodeKey]) -> GuardExpr:
    """Replace `CellRef` keys that identity-alias to another cell.

    Keys absent from `aliases`, or that already equal their image, are left
    unchanged. Composite nodes are rebuilt and interned.
    """
    if not aliases:
        return intern_guard(expr)

    def walk(node: GuardExpr) -> GuardExpr:
        if isinstance(node, CellRef):
            dest = aliases.get(node.key)
            if dest is None or dest == node.key:
                return intern_guard(node)
            return intern_guard(CellRef(dest))
        if isinstance(node, (Compare, Arith)):
            left = walk(node.left)
            right = walk(node.right)
            if left is node.left and right is node.right:
                return intern_guard(node)
            return intern_guard(type(node)(left=left, op=node.op, right=right))
        if isinstance(node, (Not, Neg)):
            operand = walk(node.operand)
            if operand is node.operand:
                return intern_guard(node)
            return intern_guard(type(node)(operand=operand))
        if isinstance(node, And):
            ops = tuple(walk(o) for o in node.operands)
            if all(a is b for a, b in zip(ops, node.operands, strict=True)):
                return intern_guard(node)
            return intern_guard(And(ops))
        if isinstance(node, Or):
            ops = tuple(walk(o) for o in node.operands)
            if all(a is b for a, b in zip(ops, node.operands, strict=True)):
                return intern_guard(node)
            return intern_guard(Or(ops))
        return intern_guard(node)

    return walk(expr)
