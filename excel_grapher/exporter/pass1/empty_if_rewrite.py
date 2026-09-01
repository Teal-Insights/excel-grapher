"""Rewrite empty-IF ``None`` arms to Excel's numeric blank (``0.0``).

excel-grapher < 3.15.3 lowered empty ``IF`` branches (``IF(cond,,x)`` /
``IF(cond,x,)``) to literal ``None``. ``xl_cell`` coerces top-level ``None``
to ``0``, so pristine oracle checks pass, but mechanically synthesized helpers
return ``None`` directly and fail batched parity.

excel-grapher 3.15.3+ emits ``0`` at the IF call site. This rewriter repairs
already-generated modules (salvage / mechanical checkpoints / exemplars that
still contain the pre-3.15.3 ``None`` arms) without requiring a full Pass 1
re-run. Only ``None`` constants that are the ``body`` or ``orelse`` of an
``ast.IfExp`` are rewritten — other ``EmptyArgNode`` uses (INDEX/MATCH) stay
as ``None``.
"""

from __future__ import annotations

import ast


class _EmptyIfNoneRewriter(ast.NodeTransformer):
    """Replace ``None`` IF-ternary arms with ``0.0``."""

    def visit_IfExp(self, node: ast.IfExp) -> ast.IfExp:
        self.generic_visit(node)
        node.body = self._rewrite_arm(node.body)
        node.orelse = self._rewrite_arm(node.orelse)
        return node

    @staticmethod
    def _rewrite_arm(arm: ast.expr) -> ast.expr:
        if isinstance(arm, ast.Constant) and arm.value is None:
            return ast.copy_location(ast.Constant(value=0.0), arm)
        return arm


def rewrite_empty_if_none_literals(source: str) -> str:
    """Return ``source`` with empty-IF ``None`` ternary arms lowered to ``0.0``.

    Idempotent: modules that already emit ``0`` / ``0.0`` are unchanged.
    """
    tree = ast.parse(source)
    rewritten = _EmptyIfNoneRewriter().visit(tree)
    ast.fix_missing_locations(rewritten)
    return ast.unparse(rewritten)
