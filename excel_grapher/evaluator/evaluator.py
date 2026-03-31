from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import TYPE_CHECKING

import fastpyxl.utils.cell

from .errors import ParseError
from .name_utils import parse_address
from .export_runtime.cache import EvalContext, xl_circular_reference, xl_iterative_compute
from .functions import FUNCTIONS
from .helpers import (
    get_error,
    to_bool,
    to_number,
    xl_column,
    xl_columns,
    xl_concat,
    xl_eq,
    xl_ge,
    xl_gt,
    xl_le,
    xl_lt,
    xl_ne,
    xl_offset_ref,
    xl_percent,
    xl_pow,
    xl_row,
)
from .parser import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
    parse,
)
from .types import CellValue, ExcelRange, XlError

_SKIP_ERROR_PRECHECK = {
    "LOOKUP",
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "XLOOKUP",
    "_XLFN.XLOOKUP",
}

if TYPE_CHECKING:
    import numpy

    from excel_grapher.grapher import DependencyGraph


@dataclass
class FormulaEvaluator:
    graph: DependencyGraph
    auto_detect_changes: bool = True
    eager_invalidation: bool = True
    on_cell_evaluated: Callable[[str, CellValue], None] | None = None
    iterate_enabled: bool = False
    iterate_count: int = 100
    iterate_delta: float = 0.001

    def __post_init__(self) -> None:
        self._cache: dict[str, CellValue] = {}
        self._call_stack: list[str] = []
        self._leaf_values: dict[str, CellValue] = {}  # For auto-detection
        self._iteration_values: dict[str, CellValue] = {}

    def __enter__(self) -> FormulaEvaluator:
        return self

    def __exit__(self, *args: object) -> None:
        return None

    def set_value(self, key: str, value: CellValue) -> None:
        """Update a cell's value and invalidate cache for it and its dependents."""
        node = self.graph.get_node(key)
        if node is None:
            raise KeyError(f"Cell {key} not found in graph")
        # Update the node's value (Node is a dataclass, so we can use object.__setattr__)
        object.__setattr__(node, "value", value)
        # Also update our tracked leaf values
        self._leaf_values[key] = value
        # Invalidate cache
        self._invalidate_with_dependents(key)

    def _invalidate_with_dependents(self, key: str) -> None:
        """Invalidate cache for a key and all cells that depend on it (transitively)."""
        to_invalidate = {key}
        # BFS to find all transitive dependents
        queue = [key]
        while queue:
            current = queue.pop(0)
            for dependent in self.graph.dependents(current):
                if dependent not in to_invalidate:
                    to_invalidate.add(dependent)
                    queue.append(dependent)
        # Remove all from cache
        for k in to_invalidate:
            self._cache.pop(k, None)

    def _iterative_target_handler(self, addr: str) -> Callable[[EvalContext, str], CellValue]:
        def handler(_ctx: EvalContext, _target: str) -> CellValue:
            return self._evaluate_cell(addr)

        return handler

    def evaluate(self, targets: list[str]) -> dict[str, CellValue]:
        # Auto-detect changes in leaf values if enabled
        if self.auto_detect_changes and self.eager_invalidation:
            self._detect_and_invalidate_changed_leaves()
        if self.iterate_enabled:
            target_handlers: dict[str, Callable[[EvalContext, str], CellValue]] = {
                addr: self._iterative_target_handler(addr) for addr in targets
            }
            ctx = EvalContext(
                inputs={},
                resolver=lambda _addr: None,
                cache=self._cache,
                computing=set(self._call_stack),
                iterative_enabled=True,
                iterate_count=self.iterate_count,
                iterate_delta=self.iterate_delta,
                iteration_values=self._iteration_values,
            )
            result = xl_iterative_compute(ctx, target_handlers)
            self._iteration_values = ctx.iteration_values
            return result
        return {addr: self._evaluate_cell(addr) for addr in targets}

    def _detect_and_invalidate_changed_leaves(self) -> None:
        """Scan all leaves and invalidate any whose values have changed."""
        for key in self.graph.leaf_keys():
            node = self.graph.get_node(key)
            if node is None:
                continue
            current_value = node.value
            if key in self._leaf_values and self._leaf_values[key] != current_value:
                self._invalidate_with_dependents(key)
            self._leaf_values[key] = current_value

    def _check_and_invalidate_if_leaves_changed(self, address: str) -> bool:
        """Check if any leaf dependencies of address have changed. Returns True if invalidated."""
        # Get all transitive dependencies (leaves) of this cell
        leaves_to_check = self._get_transitive_leaf_dependencies(address)

        changed = False
        for leaf_key in leaves_to_check:
            node = self.graph.get_node(leaf_key)
            if node is None:
                continue
            current_value = node.value
            if leaf_key in self._leaf_values and self._leaf_values[leaf_key] != current_value:
                self._invalidate_with_dependents(leaf_key)
                changed = True
            self._leaf_values[leaf_key] = current_value

        return changed

    def _get_transitive_leaf_dependencies(self, address: str) -> set[str]:
        """Get all leaf nodes that this address transitively depends on."""
        leaves: set[str] = set()
        visited: set[str] = set()
        queue = [address]

        while queue:
            current = queue.pop(0)
            if current in visited:
                continue
            visited.add(current)

            node = self.graph.get_node(current)
            if node is None:
                continue

            if node.formula is None:
                # It's a leaf
                leaves.add(current)
            else:
                # Add its dependencies to the queue
                for dep in self.graph.dependencies(current):
                    if dep not in visited:
                        queue.append(dep)

        return leaves

    def _evaluate_cell(self, address: str) -> CellValue:
        if address in self._cache:
            # Lazy invalidation: check if leaf dependencies have changed
            if self.auto_detect_changes and not self.eager_invalidation:
                if self._check_and_invalidate_if_leaves_changed(address):
                    # Cache was invalidated, need to re-evaluate (fall through)
                    pass
                else:
                    return self._cache[address]
            else:
                return self._cache[address]

        if address in self._call_stack:
            if self.iterate_enabled:
                return self._iteration_values.get(address, 0)
            return xl_circular_reference()

        node = self.graph.get_node(address)
        if node is None:
            raise KeyError(f"Cell {address} not found in graph")

        if node.formula is None:
            self._cache[address] = node.value
            self._leaf_values[address] = node.value  # Track for change detection
            if self.on_cell_evaluated is not None:
                self.on_cell_evaluated(address, node.value)
            return node.value

        formula = node.normalized_formula or node.formula
        if not isinstance(formula, str):
            raise ParseError(str(formula), "Formula is missing or not a string")

        self._call_stack.append(address)
        try:
            ast = parse(formula)
            result = self._evaluate_ast(ast)
            # Auto-resolve 1x1 ExcelRange to single value
            result = self._auto_resolve_single_cell(result)
            # Excel treats formula results of None (empty cell reference) as 0
            if result is None:
                result = 0
            self._cache[address] = result
            if self.on_cell_evaluated is not None:
                self.on_cell_evaluated(address, result)
            return result
        finally:
            self._call_stack.pop()

    def _evaluate_ast(self, node: AstNode) -> CellValue:
        if isinstance(node, EmptyArgNode):
            return None
        if isinstance(node, NumberNode):
            return node.value
        if isinstance(node, StringNode):
            return node.value
        if isinstance(node, BoolNode):
            return node.value
        if isinstance(node, ErrorNode):
            return node.error
        if isinstance(node, CellRefNode):
            return self._evaluate_cell(node.address)
        if isinstance(node, RangeNode):
            return _range_from_a1(node.start, node.end)
        if isinstance(node, FunctionCallNode):
            name = node.name.upper()
            if name == "IF":
                return self._eval_if(node.args)
            if name == "IFERROR":
                return self._eval_iferror(node.args)
            if name == "IFNA" or name == "_XLFN.IFNA":
                return self._eval_ifna(node.args)
            if name == "ISERROR":
                return self._eval_iserror(node.args)
            if name == "ISNA":
                return self._eval_isna(node.args)
            if name == "CHOOSE":
                return self._eval_choose(node.args)
            if name == "OFFSET":
                return self._eval_offset(node.args)
            if name == "ROW":
                return self._eval_row(node.args)
            if name == "COLUMN":
                return self._eval_column(node.args)
            if name == "COLUMNS":
                return self._eval_columns(node.args)

            args = [self._evaluate_ast(a) for a in node.args]
            # Resolve ExcelRange objects to numpy arrays
            args = [self._resolve_range(a) if isinstance(a, ExcelRange) else a for a in args]
            if name not in _SKIP_ERROR_PRECHECK:
                err = get_error(*args)
                if err is not None:
                    return err
            fn = FUNCTIONS.get(name)
            if fn is None:
                raise NotImplementedError(f"Excel function not implemented: {name}")
            return fn(*args)

        if isinstance(node, BinaryOpNode):
            return self._eval_binary_op(node)
        if isinstance(node, UnaryOpNode):
            return self._eval_unary_op(node)

        raise TypeError(f"Unknown AST node: {type(node)}")

    def _resolve_range(self, rng: ExcelRange) -> numpy.ndarray:
        return rng.resolve(self._evaluate_cell)

    def _auto_resolve_single_cell(self, value: CellValue) -> CellValue:
        """If value is a 1x1 ExcelRange, resolve it to its single cell value."""
        if (
            isinstance(value, ExcelRange)
            and value.start_row == value.end_row
            and value.start_col == value.end_col
        ):
            # 1x1 range - resolve to single value
            arr = self._resolve_range(value)
            return arr[0, 0]
        return value

    def _eval_binary_op(self, node: BinaryOpNode) -> CellValue:
        left = self._evaluate_ast(node.left)
        right = self._evaluate_ast(node.right)

        # Propagate errors
        if isinstance(left, XlError):
            return left
        if isinstance(right, XlError):
            return right

        op = node.op

        # String concatenation
        if op == "&":
            return xl_concat(left, right)

        # Comparison operators - handle strings case-insensitively
        if op in ("=", "<", ">", "<=", ">=", "<>"):
            cmp_fns = {
                "=": xl_eq,
                "<>": xl_ne,
                "<": xl_lt,
                ">": xl_gt,
                "<=": xl_le,
                ">=": xl_ge,
            }
            return cmp_fns[op](left, right)

        # Arithmetic operators - coerce to numbers
        ln = to_number(left)
        rn = to_number(right)
        if isinstance(ln, XlError):
            return ln
        if isinstance(rn, XlError):
            return rn

        if op == "+":
            return ln + rn
        if op == "-":
            return ln - rn
        if op == "*":
            return ln * rn
        if op == "/":
            if rn == 0:
                return XlError.DIV
            return ln / rn
        if op == "^":
            return xl_pow(left, right)

        raise ValueError(f"Unknown binary operator: {op}")

    def _eval_unary_op(self, node: UnaryOpNode) -> CellValue:
        operand = self._evaluate_ast(node.operand)
        if isinstance(operand, XlError):
            return operand

        if node.op == "-":
            n = to_number(operand)
            if isinstance(n, XlError):
                return n
            return -n

        if node.op == "%":
            return xl_percent(operand)

        raise ValueError(f"Unknown unary operator: {node.op}")

    def _eval_if(self, args: list[AstNode]) -> CellValue:
        if len(args) < 2:
            raise ParseError("IF(...)", "IF requires at least 2 arguments")
        cond = self._evaluate_ast(args[0])
        b = to_bool(cond)
        if isinstance(b, XlError):
            return b
        if b:
            return self._evaluate_ast(args[1])
        if len(args) >= 3:
            return self._evaluate_ast(args[2])
        return False

    def _eval_iferror(self, args: list[AstNode]) -> CellValue:
        if len(args) < 2:
            raise ParseError("IFERROR(...)", "IFERROR requires 2 arguments")
        v = self._evaluate_ast(args[0])
        if isinstance(v, XlError):
            return self._evaluate_ast(args[1])
        return v

    def _eval_ifna(self, args: list[AstNode]) -> CellValue:
        if len(args) < 2:
            raise ParseError("IFNA(...)", "IFNA requires 2 arguments")
        v = self._evaluate_ast(args[0])
        if v == XlError.NA:
            return self._evaluate_ast(args[1])
        return v

    def _eval_iserror(self, args: list[AstNode]) -> bool:
        if len(args) < 1:
            raise ParseError("ISERROR(...)", "ISERROR requires 1 argument")
        v = self._evaluate_ast(args[0])
        return isinstance(v, XlError)

    def _eval_isna(self, args: list[AstNode]) -> bool:
        if len(args) < 1:
            raise ParseError("ISNA(...)", "ISNA requires 1 argument")
        v = self._evaluate_ast(args[0])
        return v == XlError.NA

    def _eval_choose(self, args: list[AstNode]) -> CellValue:
        if len(args) < 2:
            raise ParseError("CHOOSE(...)", "CHOOSE requires at least 2 arguments")
        index_val = self._evaluate_ast(args[0])
        if isinstance(index_val, XlError):
            return index_val
        n = to_number(index_val)
        if isinstance(n, XlError):
            return n
        idx = int(n)
        if idx < 1 or idx > len(args) - 1:
            return XlError.VALUE
        # Only evaluate the selected choice (lazy)
        return self._evaluate_ast(args[idx])

    def _eval_offset(self, args: list[AstNode]) -> CellValue:
        if len(args) < 3:
            raise ParseError("OFFSET(...)", "OFFSET requires at least 3 arguments")

        base = self._range_from_ref_node(args[0])
        if isinstance(base, XlError):
            return base

        rows_val = self._evaluate_ast(args[1])
        cols_val = self._evaluate_ast(args[2])
        if isinstance(rows_val, XlError):
            return rows_val
        if isinstance(cols_val, XlError):
            return cols_val

        height_val = self._evaluate_ast(args[3]) if len(args) >= 4 else None
        if isinstance(height_val, XlError):
            return height_val
        width_val = self._evaluate_ast(args[4]) if len(args) >= 5 else None
        if isinstance(width_val, XlError):
            return width_val

        return xl_offset_ref(base, rows_val, cols_val, height_val, width_val)

    def _current_formula_row_col(self) -> tuple[int, int] | None:
        if not self._call_stack:
            return None
        _sheet, cell = parse_address(self._call_stack[-1])
        cell = cell.replace("$", "")
        col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
        col = fastpyxl.utils.cell.column_index_from_string(col_str)
        return row, col

    def _eval_row(self, args: list[AstNode]) -> int | XlError:
        if not args or (len(args) == 1 and isinstance(args[0], EmptyArgNode)):
            pos = self._current_formula_row_col()
            return XlError.VALUE if pos is None else pos[0]
        ref = self._range_from_ref_node(args[0])
        if isinstance(ref, XlError):
            return ref
        return xl_row(ref)

    def _eval_column(self, args: list[AstNode]) -> int | XlError:
        if not args or (len(args) == 1 and isinstance(args[0], EmptyArgNode)):
            pos = self._current_formula_row_col()
            return XlError.VALUE if pos is None else pos[1]
        ref = self._range_from_ref_node(args[0])
        if isinstance(ref, XlError):
            return ref
        return xl_column(ref)

    def _eval_columns(self, args: list[AstNode]) -> int | XlError:
        if len(args) < 1:
            raise ParseError("COLUMNS(...)", "COLUMNS requires 1 argument")
        ref = self._range_from_ref_node(args[0])
        if isinstance(ref, XlError):
            return ref
        return xl_columns(ref)

    def _range_from_ref_node(self, node: AstNode) -> ExcelRange | XlError:
        """Interpret an AST node as a reference (cell or range) without evaluating its value."""
        if isinstance(node, RangeNode):
            return _range_from_a1(node.start, node.end)

        if isinstance(node, CellRefNode):
            sheet, coord = node.address.split("!", 1)
            col_str, row = fastpyxl.utils.cell.coordinate_from_string(coord)
            col = fastpyxl.utils.cell.column_index_from_string(col_str)
            return ExcelRange(sheet=sheet, start_row=row, start_col=col, end_row=row, end_col=col)

        evaluated = self._evaluate_ast(node)
        if isinstance(evaluated, XlError):
            return evaluated
        if isinstance(evaluated, ExcelRange):
            return evaluated
        return XlError.VALUE


def _range_from_a1(start: str, end: str) -> ExcelRange:
    start_sheet, start_coord = start.split("!", 1)
    if "!" in end:
        end_sheet, end_coord = end.split("!", 1)
    else:
        end_sheet, end_coord = start_sheet, end
    if start_sheet != end_sheet:
        raise ValueError("Cross-sheet ranges are not supported")

    c1, r1 = fastpyxl.utils.cell.coordinate_from_string(start_coord)
    c2, r2 = fastpyxl.utils.cell.coordinate_from_string(end_coord)
    start_col = fastpyxl.utils.cell.column_index_from_string(c1)
    end_col = fastpyxl.utils.cell.column_index_from_string(c2)
    sr, er = sorted((r1, r2))
    sc, ec = sorted((start_col, end_col))
    return ExcelRange(sheet=start_sheet, start_row=sr, start_col=sc, end_row=er, end_col=ec)
