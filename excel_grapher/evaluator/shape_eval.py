"""Compile punched formula AST skeletons into param-taking evaluators."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import Any, cast

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    StringNode,
    UnaryOpNode,
)
from excel_grapher.core.formula_shape import (
    AddressHoleNode,
    AddressLeaf,
    SkeletonNode,
    fill_address_holes,
)
from excel_grapher.core.types import FormulaValue, XlError
from excel_grapher.evaluator.errors import ParseError
from excel_grapher.evaluator.functions import FUNCTIONS
from excel_grapher.evaluator.helpers import get_error, to_bool, to_number
from excel_grapher.evaluator.name_utils import normalize_excel_function_name

ShapeEvalFn = Callable[[tuple[AddressLeaf, ...]], FormulaValue]

_AST_SPECIAL_FUNCS = frozenset(
    {
        "IF",
        "IFERROR",
        "IFNA",
        "ISERROR",
        "ISNA",
        "ISBLANK",
        "CHOOSE",
        "OFFSET",
        "ROW",
        "COLUMN",
        "COLUMNS",
        "INDEX",
        "INDIRECT",
        "TRUE",
        "FALSE",
    }
)


def compile_formula_shape(evaluator: Any, skeleton: SkeletonNode) -> ShapeEvalFn:
    """Compile `skeleton` to a callable that binds address `params` at eval time.

    Cell/range holes dispatch to `evaluator._evaluate_ast` on the bound leaf.
    Generic functions and operators close over compiled children so a shape is
    walked/dispatched once at compile time rather than on every evaluation.
    Functions that inspect argument ASTs (OFFSET, INDEX, INDIRECT, ROW, …) fill
    holes in those arg subtrees and reuse the evaluator's existing special-case
    methods.
    """

    def compile_node(node: SkeletonNode) -> ShapeEvalFn:
        match node:
            case AddressHoleNode(_, index):
                hole_index = index

                def eval_hole(params: tuple[AddressLeaf, ...]) -> FormulaValue:
                    return evaluator._evaluate_ast(params[hole_index])

                return eval_hole
            case NumberNode(value):
                number = value

                def eval_number(_params: tuple[AddressLeaf, ...]) -> FormulaValue:
                    return number

                return eval_number
            case StringNode(value):
                text = value

                def eval_string(_params: tuple[AddressLeaf, ...]) -> FormulaValue:
                    return text

                return eval_string
            case BoolNode(value):
                flag = value

                def eval_bool(_params: tuple[AddressLeaf, ...]) -> FormulaValue:
                    return flag

                return eval_bool
            case ErrorNode(error):
                err = error

                def eval_error(_params: tuple[AddressLeaf, ...]) -> FormulaValue:
                    return err

                return eval_error
            case EmptyArgNode():

                def eval_empty(_params: tuple[AddressLeaf, ...]) -> FormulaValue:
                    return None

                return eval_empty
            case BinaryOpNode(op, left, right):
                left_fn = compile_node(cast(SkeletonNode, left))
                right_fn = compile_node(cast(SkeletonNode, right))
                operator = op

                def eval_binary(params: tuple[AddressLeaf, ...]) -> FormulaValue:
                    left_v = evaluator._resolve_binary_operand(left_fn(params))
                    right_v = evaluator._resolve_binary_operand(right_fn(params))
                    return evaluator._apply_binary_op(operator, left_v, right_v)

                return eval_binary
            case UnaryOpNode(op, operand):
                operand_fn = compile_node(cast(SkeletonNode, operand))
                operator = op

                def eval_unary(params: tuple[AddressLeaf, ...]) -> FormulaValue:
                    value = evaluator._resolve_binary_operand(operand_fn(params))
                    return evaluator._apply_unary_op(operator, value)

                return eval_unary
            case FunctionCallNode(name, args):
                return _compile_function(evaluator, name, args, compile_node)
        raise TypeError(f"unsupported skeleton node: {type(node).__name__}")

    return compile_node(skeleton)


def _compile_function(
    evaluator: Any,
    raw_name: str,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    name = normalize_excel_function_name(raw_name)
    if name in _AST_SPECIAL_FUNCS:
        return _compile_special_function(evaluator, name, args, compile_node)

    arg_fns = [compile_node(cast(SkeletonNode, arg)) for arg in args]
    from excel_grapher.evaluator.evaluator import _SKIP_ERROR_PRECHECK

    skip_precheck = name in _SKIP_ERROR_PRECHECK
    fn = FUNCTIONS.get(name)

    def eval_generic(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        values = [arg_fn(params) for arg_fn in arg_fns]
        values = [
            evaluator._resolve_function_arg(value, name, index)
            for index, value in enumerate(values)
        ]
        if not skip_precheck:
            err = get_error(*values)
            if err is not None:
                return err
        if fn is None:
            raise NotImplementedError(f"Excel function not implemented: {name}")
        return fn(*values)

    return eval_generic


def _compile_special_function(
    evaluator: Any,
    name: str,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    if name == "TRUE":

        def eval_true(_params: tuple[AddressLeaf, ...]) -> FormulaValue:
            return True

        return eval_true
    if name == "FALSE":

        def eval_false(_params: tuple[AddressLeaf, ...]) -> FormulaValue:
            return False

        return eval_false
    if name == "IF":
        return _compile_if(evaluator, args, compile_node)
    if name == "IFERROR":
        return _compile_iferror(evaluator, args, compile_node)
    if name == "IFNA":
        return _compile_ifna(evaluator, args, compile_node)
    if name == "ISERROR":
        return _compile_iserror(evaluator, args, compile_node)
    if name == "ISNA":
        return _compile_isna(evaluator, args, compile_node)
    if name == "ISBLANK":
        return _compile_isblank(evaluator, args, compile_node)
    if name == "CHOOSE":
        return _compile_choose(evaluator, args, compile_node)

    # OFFSET / INDEX / INDIRECT / ROW / COLUMN / COLUMNS inspect argument ASTs.
    skeletons = tuple(cast(SkeletonNode, arg) for arg in args)

    def eval_ref_fn(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        filled = [fill_address_holes(arg, params) for arg in skeletons]
        if name == "OFFSET":
            return evaluator._eval_offset(filled)
        if name == "INDEX":
            return evaluator._eval_index(filled)
        if name == "INDIRECT":
            return evaluator._eval_indirect(filled)
        if name == "ROW":
            return evaluator._eval_row(filled)
        if name == "COLUMN":
            return evaluator._eval_column(filled)
        if name == "COLUMNS":
            return evaluator._eval_columns(filled)
        raise AssertionError(f"unhandled special function: {name}")

    return eval_ref_fn


def _compile_if(
    _evaluator: Any,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    if len(args) < 2:
        raise ParseError("IF(...)", "IF requires at least 2 arguments")
    cond_fn = compile_node(cast(SkeletonNode, args[0]))
    then_fn = compile_node(cast(SkeletonNode, args[1]))
    else_fn = compile_node(cast(SkeletonNode, args[2])) if len(args) >= 3 else None
    then_empty = isinstance(args[1], EmptyArgNode)
    else_empty = len(args) >= 3 and isinstance(args[2], EmptyArgNode)

    def eval_if(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        cond = cond_fn(params)
        flag = to_bool(cond)
        if isinstance(flag, XlError):
            return flag
        if flag:
            return 0 if then_empty else then_fn(params)
        if else_fn is None:
            return False
        return 0 if else_empty else else_fn(params)

    return eval_if


def _compile_iferror(
    _evaluator: Any,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    if len(args) < 2:
        raise ParseError("IFERROR(...)", "IFERROR requires 2 arguments")
    value_fn = compile_node(cast(SkeletonNode, args[0]))
    fallback_fn = compile_node(cast(SkeletonNode, args[1]))

    def eval_iferror(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        value = value_fn(params)
        if isinstance(value, XlError):
            return fallback_fn(params)
        return value

    return eval_iferror


def _compile_ifna(
    _evaluator: Any,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    if len(args) < 2:
        raise ParseError("IFNA(...)", "IFNA requires 2 arguments")
    value_fn = compile_node(cast(SkeletonNode, args[0]))
    fallback_fn = compile_node(cast(SkeletonNode, args[1]))

    def eval_ifna(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        value = value_fn(params)
        if value == XlError.NA:
            return fallback_fn(params)
        return value

    return eval_ifna


def _compile_iserror(
    _evaluator: Any,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    if len(args) < 1:
        raise ParseError("ISERROR(...)", "ISERROR requires 1 argument")
    value_fn = compile_node(cast(SkeletonNode, args[0]))

    def eval_iserror(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        return isinstance(value_fn(params), XlError)

    return eval_iserror


def _compile_isna(
    _evaluator: Any,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    if len(args) < 1:
        raise ParseError("ISNA(...)", "ISNA requires 1 argument")
    value_fn = compile_node(cast(SkeletonNode, args[0]))

    def eval_isna(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        return value_fn(params) == XlError.NA

    return eval_isna


def _compile_isblank(
    evaluator: Any,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    del compile_node
    if len(args) != 1:
        raise ParseError("ISBLANK(...)", "ISBLANK requires 1 argument")
    skeleton = cast(SkeletonNode, args[0])

    def eval_isblank(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        return evaluator._eval_isblank([fill_address_holes(skeleton, params)])

    return eval_isblank


def _compile_choose(
    _evaluator: Any,
    args: Sequence[AstNode],
    compile_node: Callable[[SkeletonNode], ShapeEvalFn],
) -> ShapeEvalFn:
    if len(args) < 2:
        raise ParseError("CHOOSE(...)", "CHOOSE requires at least 2 arguments")
    index_fn = compile_node(cast(SkeletonNode, args[0]))
    choice_fns = [compile_node(cast(SkeletonNode, arg)) for arg in args[1:]]

    def eval_choose(params: tuple[AddressLeaf, ...]) -> FormulaValue:
        index_val = index_fn(params)
        if isinstance(index_val, XlError):
            return index_val
        number = to_number(index_val)
        if isinstance(number, XlError):
            return number
        idx = int(number)
        if idx < 1 or idx > len(choice_fns):
            return XlError.VALUE
        return choice_fns[idx - 1](params)

    return eval_choose
