"""Hoist nested calls in generated cell return expressions into temporaries."""

from __future__ import annotations

import ast

__all__ = ["unpack_return_expression"]

_NON_HOISTABLE_CALLS = frozenset(
    {
        "xl_number",
        "xl_bool",
        "xl_int",
        "xl_raise",
        "to_string",
    }
)


def _is_hoistable_call(node: ast.Call) -> bool:
    if isinstance(node.func, ast.Name):
        name = node.func.id
        if name in _NON_HOISTABLE_CALLS:
            return False
        if name.startswith("xl_") or name == "ExcelRange":
            return True
    return False


def unpack_return_expression(expr: str, temp_start: int) -> tuple[list[str], str, int]:
    """Unpack nested function calls in a return expression into statement temporaries.

    Nested calls are assigned to ``_tN`` variables in post-order (call order) and
    the outermost call remains in the returned expression. Lazy regions (lambda
    bodies and conditional-expression branches) are left unchanged so Excel-style
    short-circuit semantics are preserved.

    Args:
        expr: A single Python expression string emitted for a cell ``return``.
        temp_start: Number of ``_tN`` names already used while emitting ``expr``.

    Returns:
        A tuple of statement lines (without indentation), the return expression,
        and the updated temporary counter (last ``N`` used in ``_tN`` names).
    """
    tree = ast.parse(expr, mode="eval")
    statements: list[str] = []
    counter = temp_start

    def _next_name() -> str:
        nonlocal counter
        counter += 1
        return f"_t{counter}"

    def _transform(
        node: ast.expr,
        *,
        is_root: bool,
        lazy: bool,
        parent_is_call: bool,
    ) -> ast.expr:
        if lazy:
            return node

        if isinstance(node, ast.Lambda):
            return node

        if isinstance(node, ast.IfExp):
            return ast.IfExp(
                test=_transform(
                    node.test,
                    is_root=False,
                    lazy=False,
                    parent_is_call=False,
                ),
                body=_transform(node.body, is_root=False, lazy=True, parent_is_call=False),
                orelse=_transform(node.orelse, is_root=False, lazy=True, parent_is_call=False),
            )

        if isinstance(node, ast.NamedExpr):
            return ast.NamedExpr(
                target=node.target,
                value=_transform(
                    node.value,
                    is_root=False,
                    lazy=False,
                    parent_is_call=False,
                ),
            )

        if isinstance(node, ast.Call):
            func = _transform(
                node.func,
                is_root=False,
                lazy=False,
                parent_is_call=False,
            )
            had_call_arg = any(isinstance(arg, ast.Call) for arg in node.args)
            args = [
                _transform(arg, is_root=False, lazy=False, parent_is_call=True) for arg in node.args
            ]
            keywords = [
                ast.keyword(
                    kw.arg,
                    _transform(kw.value, is_root=False, lazy=False, parent_is_call=True),
                )
                for kw in node.keywords
            ]
            new_call = ast.Call(func=func, args=args, keywords=keywords)
            ast.fix_missing_locations(new_call)
            if is_root:
                return new_call

            if not _is_hoistable_call(node):
                return new_call

            if had_call_arg or parent_is_call:
                name = _next_name()
                statements.append(f"{name} = {ast.unparse(new_call)}")
                return ast.Name(id=name, ctx=ast.Load())

            return new_call

        if isinstance(node, ast.UnaryOp):
            return ast.UnaryOp(
                op=node.op,
                operand=_transform(
                    node.operand,
                    is_root=False,
                    lazy=False,
                    parent_is_call=parent_is_call,
                ),
            )

        if isinstance(node, ast.BinOp):
            return ast.BinOp(
                left=_transform(
                    node.left,
                    is_root=False,
                    lazy=False,
                    parent_is_call=parent_is_call,
                ),
                op=node.op,
                right=_transform(
                    node.right,
                    is_root=False,
                    lazy=False,
                    parent_is_call=parent_is_call,
                ),
            )

        if isinstance(node, ast.BoolOp):
            return ast.BoolOp(
                op=node.op,
                values=[
                    _transform(value, is_root=False, lazy=True, parent_is_call=False)
                    for value in node.values
                ],
            )

        if isinstance(node, ast.Compare):
            return ast.Compare(
                left=_transform(
                    node.left,
                    is_root=False,
                    lazy=False,
                    parent_is_call=parent_is_call,
                ),
                ops=node.ops,
                comparators=[
                    _transform(comp, is_root=False, lazy=False, parent_is_call=parent_is_call)
                    for comp in node.comparators
                ],
            )

        if isinstance(node, ast.Subscript):
            return ast.Subscript(
                value=_transform(
                    node.value,
                    is_root=False,
                    lazy=False,
                    parent_is_call=parent_is_call,
                ),
                slice=_transform(
                    node.slice,
                    is_root=False,
                    lazy=False,
                    parent_is_call=parent_is_call,
                ),
                ctx=node.ctx,
            )

        if isinstance(node, ast.Attribute):
            return ast.Attribute(
                value=_transform(
                    node.value,
                    is_root=False,
                    lazy=False,
                    parent_is_call=parent_is_call,
                ),
                attr=node.attr,
                ctx=node.ctx,
            )

        if isinstance(node, (ast.Tuple, ast.List)):
            return type(node)(
                elts=[
                    _transform(elt, is_root=False, lazy=False, parent_is_call=parent_is_call)
                    for elt in node.elts
                ],
                ctx=node.ctx,
            )

        if isinstance(node, ast.Starred):
            return ast.Starred(
                value=_transform(
                    node.value,
                    is_root=False,
                    lazy=False,
                    parent_is_call=parent_is_call,
                ),
                ctx=node.ctx,
            )

        if isinstance(node, ast.Dict):
            return ast.Dict(
                keys=[
                    None
                    if key is None
                    else _transform(key, is_root=False, lazy=False, parent_is_call=parent_is_call)
                    for key in node.keys
                ],
                values=[
                    _transform(value, is_root=False, lazy=False, parent_is_call=parent_is_call)
                    for value in node.values
                ],
            )

        if isinstance(node, ast.Set):
            return ast.Set(
                elts=[
                    _transform(elt, is_root=False, lazy=False, parent_is_call=parent_is_call)
                    for elt in node.elts
                ],
            )

        return node

    new_body = _transform(tree.body, is_root=True, lazy=False, parent_is_call=False)
    ast.fix_missing_locations(new_body)
    return statements, ast.unparse(new_body), counter
