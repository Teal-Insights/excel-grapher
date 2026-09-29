"""Helpers for tightening may-cycle feasibility (#533)."""

from __future__ import annotations

import math
from collections.abc import Mapping
from dataclasses import dataclass

from excel_grapher.core.cell_types import CellTypeEnv
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    UnaryOpNode,
    resolve_cell_ref,
)

from .guard import And, CellRef, Compare, GuardExpr, Literal, Not, Or, _constraint_key, intern_guard
from .node import Node, NodeKey


def identity_alias_map(nodes: Mapping[NodeKey, Node]) -> dict[NodeKey, NodeKey]:
    """Map identity-formula cells to the cell they copy, followed to a root.

    A node is an identity alias when `formula_ast` is a single `CellRefNode`.
    Relative axes resolve against the host cell. Cycles in the alias graph stop
    at the first repeat (the repeating node is the root).
    """
    parent: dict[NodeKey, NodeKey] = {}
    for key, node in nodes.items():
        ast = node.formula_ast
        if not isinstance(ast, CellRefNode):
            continue
        try:
            target = resolve_cell_ref(ast, key)
        except ValueError:
            continue
        parent[key] = target

    def root(key: NodeKey) -> NodeKey:
        seen: set[NodeKey] = set()
        while key in parent and key not in seen:
            seen.add(key)
            key = parent[key]
        return key

    return {key: root(key) for key in parent if root(key) != key}


# ---- affine / interval abstraction of guard-cone cells (#1028) ---------------

# Absolute slack for comparisons that assert a strict relation, so float noise
# (and Excel's near-equality) never turns "maybe equal" into "definitely not".
_TOL = 1e-9


@dataclass(frozen=True, slots=True)
class AbstractValue:
    """Sound over-approximation of a cell's value for deciding guard atoms.

    Attributes:
        root: When set, the value is exactly `root + offset`; `None` when only
            the bounds are known.
        offset: Additive offset from `root`.
        lo: Lower bound on the value (`-inf` when unknown).
        hi: Upper bound on the value (`inf` when unknown).
        numeric: True when the value is a number or an error. False when it may
            also be text, a boolean, or blank, which Excel orders differently.
    """

    root: NodeKey | None
    offset: float
    lo: float
    hi: float
    numeric: bool

    @property
    def is_const(self) -> bool:
        """Return whether the value is a single known number."""
        return self.numeric and self.lo == self.hi


def _const(value: float) -> AbstractValue:
    return AbstractValue(None, 0.0, value, value, True)


def _range(lo: float, hi: float) -> AbstractValue:
    return AbstractValue(None, 0.0, lo, hi, True)


def _shift(a: AbstractValue, delta: float) -> AbstractValue:
    return AbstractValue(a.root, a.offset + delta, a.lo + delta, a.hi + delta, True)


def _add(a: AbstractValue, b: AbstractValue) -> AbstractValue:
    if b.is_const:
        return _shift(a, b.lo)
    if a.is_const:
        return _shift(b, a.lo)
    return _range(a.lo + b.lo, a.hi + b.hi)


def _sub(a: AbstractValue, b: AbstractValue) -> AbstractValue:
    if a.root is not None and a.root == b.root:
        return _const(a.offset - b.offset)
    if b.is_const:
        return _shift(a, -b.lo)
    return _range(a.lo - b.hi, a.hi - b.lo)


def _extreme(args: list[AbstractValue], *, upper: bool) -> AbstractValue | None:
    """Abstract `MAX(args)` (`upper`) or `MIN(args)` over numeric operands."""
    if not args or not all(a.numeric for a in args):
        return None
    pick = max if upper else min
    roots = {a.root for a in args}
    if len(roots) == 1 and None not in roots:
        return max(args, key=lambda a: a.offset if upper else -a.offset)
    return _range(pick(a.lo for a in args), pick(a.hi for a in args))


def _env_value(env: CellTypeEnv | None, key: NodeKey) -> AbstractValue:
    """Opaque value of `key`, bounded by its numeric domain when one is declared."""
    cell_type = None if env is None else env.get(_constraint_key(key))
    if cell_type is None:
        return AbstractValue(key, 0.0, -math.inf, math.inf, False)
    parts: list[tuple[float, float]] = []
    if cell_type.enum is not None:
        values = cell_type.enum.values
        if not values or any(
            isinstance(v, bool) or not isinstance(v, (int, float)) for v in values
        ):
            return AbstractValue(key, 0.0, -math.inf, math.inf, False)
        parts.append((float(min(values)), float(max(values))))
    for dom in (cell_type.interval, cell_type.real_interval):
        if dom is not None:
            lo = -math.inf if dom.min is None else float(dom.min)
            hi = math.inf if dom.max is None else float(dom.max)
            parts.append((lo, hi))
    if not parts:
        return AbstractValue(key, 0.0, -math.inf, math.inf, False)
    # Union domains are a hull; intersected domains stay sound as a hull too.
    return AbstractValue(key, 0.0, min(lo for lo, _ in parts), max(hi for _, hi in parts), True)


def _definitely_lt(a: AbstractValue, b: AbstractValue) -> bool:
    return a.hi < b.lo - _TOL


def _definitely_le(a: AbstractValue, b: AbstractValue) -> bool:
    return a.hi <= b.lo


def _decide(op: str, a: AbstractValue, b: AbstractValue) -> bool | None:
    """Return the truth of `a op b` when it holds for every concrete value."""
    if a.root is not None and a.root == b.root:
        d = a.offset - b.offset
        if op in ("=", "<>"):
            # Unequal offsets mean at least one side is arithmetic (numeric or
            # an error), so the sides can never compare equal.
            if d == 0:
                return op == "="
            if abs(d) > _TOL:
                return op == "<>"
            return None
        if not (a.numeric and b.numeric):
            return None
        a, b = _const(d), _const(0.0)
    if not (a.numeric and b.numeric):
        return None
    if op == "=":
        if _definitely_lt(a, b) or _definitely_lt(b, a):
            return False
        return True if _definitely_le(a, b) and _definitely_le(b, a) else None
    if op == "<>":
        out = _decide("=", a, b)
        return None if out is None else not out
    if op in (">", ">="):
        a, b = b, a
        op = "<" if op == ">" else "<="
    if op == "<":
        if _definitely_lt(a, b):
            return True
        return False if _definitely_le(b, a) else None
    if op == "<=":
        if _definitely_le(a, b):
            return True
        return False if _definitely_lt(b, a) else None
    return None


class GuardConeAbstraction:
    """Lazily abstract guard cells as `root + offset` or numeric intervals.

    Recognised formula shapes are numeric literals, `=R`, `R+n`, `R-n`,
    `a+b`, `a-b`, `MAX(...)`, `MIN(...)` and `IF(x>y,x,y)` style max/min.
    Every other cell is its own opaque root, bounded only by its declared
    `CellTypeEnv` domain, which keeps unknown shapes sound.

    Args:
        nodes: Graph nodes keyed by cell key.
        cell_type_env: Leaf domains used to bound opaque roots.
    """

    def __init__(self, nodes: Mapping[NodeKey, Node], cell_type_env: CellTypeEnv | None) -> None:
        self._nodes = nodes
        self._env = cell_type_env
        self._memo: dict[NodeKey, AbstractValue] = {}
        # Guards are interned, so identity keys hit across repeated DFS visits.
        self._folded: dict[int, tuple[GuardExpr, GuardExpr | bool]] = {}

    def value(self, key: NodeKey) -> AbstractValue:
        """Return the abstract value of cell `key`."""
        if key in self._memo:
            return self._memo[key]
        # Iterative post-order so long `=R+1` chains do not hit the recursion limit.
        stack = [key]
        in_progress: set[NodeKey] = set()
        while stack:
            cur = stack[-1]
            if cur in self._memo:
                stack.pop()
                continue
            refs = self._shape_refs(cur)
            pending = [r for r in refs if r not in self._memo] if refs is not None else []
            if cur not in in_progress and pending:
                in_progress.add(cur)
                # A reference back into an in-progress cell is a cycle: stay opaque.
                if any(r in in_progress for r in pending):
                    self._memo[cur] = _env_value(self._env, cur)
                    stack.pop()
                    continue
                stack.extend(pending)
                continue
            stack.pop()
            in_progress.discard(cur)
            if refs is None or pending:
                self._memo[cur] = _env_value(self._env, cur)
                continue
            ast = self._nodes[cur].formula_ast
            out = None if ast is None else self._eval(ast, cur)
            self._memo[cur] = _env_value(self._env, cur) if out is None else out
        return self._memo[key]

    def _shape_refs(self, key: NodeKey) -> list[NodeKey] | None:
        """Return cell refs of a recognised formula shape, or `None` for opaque cells."""
        node = self._nodes.get(key)
        if node is None or node.formula_ast is None:
            return None
        refs: list[NodeKey] = []

        def walk(ast: AstNode) -> bool:
            if isinstance(ast, NumberNode):
                return True
            if isinstance(ast, CellRefNode):
                try:
                    refs.append(resolve_cell_ref(ast, key))
                except ValueError:
                    return False
                return True
            if isinstance(ast, BinaryOpNode) and ast.op in ("+", "-", ">", ">=", "<", "<="):
                return walk(ast.left) and walk(ast.right)
            if isinstance(ast, UnaryOpNode) and ast.op == "-":
                return walk(ast.operand)
            if isinstance(ast, FunctionCallNode) and ast.name.upper() in ("MAX", "MIN", "IF"):
                return bool(ast.args) and all(walk(a) for a in ast.args)
            return False

        return refs if walk(node.formula_ast) else None

    def _eval(self, ast: AstNode, host: NodeKey) -> AbstractValue | None:
        if isinstance(ast, NumberNode):
            return _const(float(ast.value))
        if isinstance(ast, CellRefNode):
            return self._memo[resolve_cell_ref(ast, host)]
        if isinstance(ast, BinaryOpNode) and ast.op in ("+", "-"):
            a, b = self._eval(ast.left, host), self._eval(ast.right, host)
            if a is None or b is None:
                return None
            return _add(a, b) if ast.op == "+" else _sub(a, b)
        if isinstance(ast, UnaryOpNode) and ast.op == "-":
            a = self._eval(ast.operand, host)
            return None if a is None else _range(-a.hi, -a.lo)
        if isinstance(ast, FunctionCallNode):
            name = ast.name.upper()
            if name in ("MAX", "MIN"):
                args = [self._eval(a, host) for a in ast.args]
                if any(a is None for a in args):
                    return None
                return _extreme([a for a in args if a is not None], upper=name == "MAX")
            if name == "IF":
                return self._eval_if_extreme(ast, host)
        return None

    def _eval_if_extreme(self, ast: FunctionCallNode, host: NodeKey) -> AbstractValue | None:
        """Abstract `IF(x>y,x,y)` as `MAX(x,y)` and `IF(x<y,x,y)` as `MIN(x,y)`."""
        if len(ast.args) != 3:
            return None
        cond, then, other = ast.args
        if not isinstance(cond, BinaryOpNode) or cond.op not in (">", ">=", "<", "<="):
            return None
        if (cond.left, cond.right) == (then, other):
            upper = cond.op in (">", ">=")
        elif (cond.left, cond.right) == (other, then):
            upper = cond.op in ("<", "<=")
        else:
            return None
        x, y = self._eval(then, host), self._eval(other, host)
        if x is None or y is None:
            return None
        return _extreme([x, y], upper=upper)

    def _term(self, expr: GuardExpr) -> AbstractValue | None:
        if isinstance(expr, CellRef):
            return self.value(expr.key)
        if isinstance(expr, Literal):
            v = expr.value
            if isinstance(v, bool) or not isinstance(v, (int, float)):
                return None
            return _const(float(v))
        return None

    def simplify(self, guard: GuardExpr) -> GuardExpr | bool:
        """Fold `Compare` atoms that the abstraction decides.

        Returns:
            `True` or `False` when the whole guard is decided, else the guard
            with decided atoms removed (undecided atoms are left as-is).
        """
        hit = self._folded.get(id(guard))
        if hit is not None and hit[0] is guard:
            return hit[1]
        out = self._simplify(guard)
        self._folded[id(guard)] = (guard, out)
        return out

    def _simplify(self, guard: GuardExpr) -> GuardExpr | bool:
        if isinstance(guard, Compare):
            a, b = self._term(guard.left), self._term(guard.right)
            if a is None or b is None:
                return guard
            out = _decide(guard.op, a, b)
            return guard if out is None else out
        if isinstance(guard, Not):
            inner = self._simplify(guard.operand)
            if isinstance(inner, bool):
                return not inner
            return guard if inner is guard.operand else intern_guard(Not(inner))
        if isinstance(guard, (And, Or)):
            absorbing = isinstance(guard, Or)
            kept: list[GuardExpr] = []
            for operand in guard.operands:
                out = self._simplify(operand)
                if out is absorbing:
                    return absorbing
                if not isinstance(out, bool):
                    kept.append(out)
            if not kept:
                return not absorbing
            if len(kept) == 1:
                return kept[0]
            cls = Or if absorbing else And
            return intern_guard(cls(tuple(kept)))
        return guard
