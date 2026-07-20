from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, cast, overload

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    CellKey,
    parse_address,
    parse_node_key,
)
from excel_grapher.core.address_keys import (
    normalize_key as normalize_address,
)
from excel_grapher.core.addressing import index_excel_range
from excel_grapher.core.excel_function_meta import (
    eager_materialize_arg_indices,
    grid_range_arg_indices,
)
from excel_grapher.core.grid import Range
from excel_grapher.core.range_shorthand import (
    SheetBounds,
    resolve_whole_column,
    resolve_whole_row,
)
from excel_grapher.core.types import CellValue, ExcelRange, XlError, resolve_excel_range
from excel_grapher.evaluator.name_utils import normalize_excel_function_name
from excel_grapher.grapher.blank_ranges import (
    address_in_blank_ranges,
    normalize_blank_range_specs,
)
from excel_grapher.grapher.formula_groups import specialize_group
from excel_grapher.grapher.node import NodeKind, NodeView, locate_cell
from excel_grapher.runtime.cache import (
    EvalContext,
    warn_circular_reference,
    xl_circular_reference,
    xl_iterative_compute,
)
from excel_grapher.runtime.info import xl_isblank

from .ast_cache import DEFAULT_AST_CACHE_MAXSIZE, AstCache, AstCacheInfo
from .errors import (
    FormulaGroupKeyError,
    MissingGroupTemplateError,
    MissingNormalizedFormulaError,
    ParseError,
)
from .functions import FUNCTIONS
from .helpers import (
    get_error,
    to_bool,
    to_number,
    xl_add,
    xl_column,
    xl_columns,
    xl_concat,
    xl_div,
    xl_eq,
    xl_ge,
    xl_gt,
    xl_le,
    xl_lt,
    xl_mul,
    xl_ne,
    xl_offset_ref,
    xl_percent,
    xl_pow,
    xl_row,
    xl_sub,
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
    WholeColumnNode,
    WholeRowNode,
    parse,
)

_SKIP_ERROR_PRECHECK = {
    "LOOKUP",
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "XLOOKUP",
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
    blank_ranges: tuple[str, ...] | None = None
    ast_cache_maxsize: int = DEFAULT_AST_CACHE_MAXSIZE

    def __post_init__(self) -> None:
        self._cache: dict[str, CellValue] = {}
        self._ast_cache = AstCache(maxsize=self.ast_cache_maxsize)
        preparsed = self.graph.preparsed_formulas
        if preparsed:
            self._ast_cache.seed(preparsed)
        self._call_stack: list[str] = []
        self._leaf_values: dict[str, CellValue] = {}  # For auto-detection
        self._iteration_values: dict[str, CellValue] = {}
        self._blank_range_rects = normalize_blank_range_specs(self.blank_ranges)
        # Runtime dependency edges recorded as cells are evaluated. Unlike the
        # static graph edges (which freeze the build-time resolution of dynamic
        # refs such as OFFSET/INDIRECT under `use_cached_dynamic_refs`), these
        # follow argument-driven resolution shifts, so invalidation tracks the
        # currently-resolved dependency chain. This mirrors the exported
        # `EvalContext` runtime, keeping evaluator and export in parity.
        self._runtime_deps: dict[str, set[str]] = {}
        self._runtime_reverse_deps: dict[str, set[str]] = {}
        self._circular_warning_roots: set[str] = set()

    def __enter__(self) -> FormulaEvaluator:
        return self

    def __exit__(self, *args: object) -> None:
        return None

    def clear_caches(self) -> None:
        """Clear cached cell values and parsed formula ASTs."""
        self._cache.clear()
        self._circular_warning_roots.clear()
        self._ast_cache.clear()

    def ast_cache_info(self) -> AstCacheInfo:
        """Return hit/miss statistics for the AST parse cache."""
        return self._ast_cache.cache_info()

    def _parse_cached(self, normalized_formula: str) -> AstNode:
        return self._ast_cache.get(normalized_formula, parse_fn=parse)

    def _record_runtime_dependency(self, parent: str, child: str) -> None:
        """Record that `parent` read `child` during the current evaluation."""
        if parent == child:
            return
        self._runtime_deps.setdefault(parent, set()).add(child)
        self._runtime_reverse_deps.setdefault(child, set()).add(parent)

    def _invalidate_with_dependents(self, key: str) -> None:
        """Invalidate cache for a key and all cells that depend on it (transitively).

        Walks runtime dependency edges recorded during evaluation rather than the
        static graph edges, so dependents reached through a shifted dynamic-ref
        resolution are invalidated. Invalidated cells drop their outgoing edges so
        the current dependency chain is re-recorded on recompute.
        """
        to_visit = [key]
        seen: set[str] = set()
        while to_visit:
            addr = to_visit.pop()
            if addr in seen:
                continue
            seen.add(addr)
            self._cache.pop(addr, None)
            self._circular_warning_roots.discard(addr)
            to_visit.extend(self._runtime_reverse_deps.get(addr, set()))
            for dep in self._runtime_deps.get(addr, set()):
                parents = self._runtime_reverse_deps.get(dep)
                if parents is not None:
                    parents.discard(addr)
                    if not parents:
                        self._runtime_reverse_deps.pop(dep, None)
            self._runtime_deps.pop(addr, None)
            self._runtime_reverse_deps.pop(addr, None)

    def _iterative_target_handler(self, addr: str) -> Callable[[EvalContext, str], CellValue]:
        def handler(_ctx: EvalContext, _target: str) -> CellValue:
            return self._evaluate_cell(addr)

        return handler

    @overload
    def evaluate(self, targets: str) -> CellValue: ...

    @overload
    def evaluate(self, targets: Sequence[str]) -> dict[str, CellValue]: ...

    def evaluate(self, targets: str | Sequence[str]) -> CellValue | dict[str, CellValue]:
        if isinstance(targets, str):
            single = True
            target_list: list[str] = [targets]
        else:
            single = False
            target_list = list(targets)
        # Auto-detect changes in leaf values if enabled
        if self.auto_detect_changes and self.eager_invalidation:
            self._detect_and_invalidate_changed_leaves()
        if self.iterate_enabled:
            target_handlers: dict[str, Callable[[EvalContext, str], CellValue]] = {
                addr: self._iterative_target_handler(addr) for addr in target_list
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
            return next(iter(result.values())) if single else result
        results = {addr: self._evaluate_cell(addr) for addr in target_list}
        return next(iter(results.values())) if single else results

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
        """Get all leaf nodes that this address transitively depends on.

        Walks runtime dependency edges (recorded during evaluation) so the leaf
        set reflects the currently-resolved dynamic-ref chain rather than the
        static build-time resolution.
        """
        leaves: set[str] = set()
        visited: set[str] = set()
        queue = [normalize_address(address)]

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
                for dep in self._runtime_deps.get(current, set()):
                    if dep not in visited:
                        queue.append(dep)

        return leaves

    def _evaluate_cell(self, address: str) -> CellValue:
        try:
            parsed = parse_node_key(address)
        except ValueError:
            parsed = None
            norm = normalize_address(address)
        else:
            norm = str(parsed)

        # Public API is member/cell addresses only (unique occupancy).
        if parsed is not None and not isinstance(parsed, CellKey):
            raise FormulaGroupKeyError(norm)
        if parsed is None:
            try:
                reparsed = parse_node_key(norm)
            except ValueError:
                reparsed = None
            if reparsed is not None and not isinstance(reparsed, CellKey):
                raise FormulaGroupKeyError(norm)
            if reparsed is not None:
                norm = str(reparsed)

        if self._call_stack:
            self._record_runtime_dependency(self._call_stack[-1], norm)
        if norm in self._cache:
            # Lazy invalidation: check if leaf dependencies have changed
            if self.auto_detect_changes and not self.eager_invalidation:
                if self._check_and_invalidate_if_leaves_changed(norm):
                    # Cache was invalidated, need to re-evaluate (fall through)
                    pass
                else:
                    if norm in self._circular_warning_roots:
                        warn_circular_reference(stacklevel=3)
                    return self._cache[norm]
            else:
                if norm in self._circular_warning_roots:
                    warn_circular_reference(stacklevel=3)
                return self._cache[norm]

        if norm in self._call_stack:
            if self.iterate_enabled:
                return self._iteration_values.get(norm, 0)
            root = self._call_stack[0]
            self._circular_warning_roots.add(root)
            return xl_circular_reference()

        if self._blank_range_rects and address_in_blank_ranges(norm, self._blank_range_rects):
            self._cache[norm] = None
            if self.on_cell_evaluated is not None:
                self.on_cell_evaluated(norm, None)
            return None

        location = locate_cell(self.graph, norm)
        if location is None:
            raise KeyError(f"Cell {address} not found in graph")

        node = self.graph.get_node(location.node_key)
        if node is None:
            raise KeyError(f"Cell {address} not found in graph")

        if location.kind is not NodeKind.cell:
            return self._evaluate_group_member(norm, location.node_key, node)

        if node.formula is None:
            self._cache[norm] = node.value
            self._leaf_values[norm] = node.value  # Track for change detection
            if self.on_cell_evaluated is not None:
                self.on_cell_evaluated(norm, node.value)
            return node.value

        nf = node.normalized_formula
        if nf is None or not isinstance(nf, str) or not nf.strip():
            raise MissingNormalizedFormulaError(norm)
        formula = nf.strip()

        self._call_stack.append(norm)
        try:
            ast = self._parse_cached(formula)
            result = self._evaluate_ast(ast)
            # Auto-resolve 1x1 ExcelRange to single value
            result = self._auto_resolve_single_cell(result)
            # Excel treats formula results of None (empty cell reference) as 0
            if result is None:
                result = 0
            self._cache[norm] = result
            if self.on_cell_evaluated is not None:
                self.on_cell_evaluated(norm, result)
            return result
        finally:
            self._call_stack.pop()

    def _evaluate_group_member(self, member_key: str, group_key: str, node: NodeView) -> CellValue:
        """Specialize a formula-group template for `member_key` and evaluate it.

        Results are cached under the **member** address so sibling members stay
        lazy.
        """
        if node.skeleton is None or node.member_bindings is None:
            raise MissingGroupTemplateError(group_key, member_key)
        bindings = node.member_bindings.get(member_key)
        if bindings is None:
            raise MissingGroupTemplateError(group_key, member_key)

        self._call_stack.append(member_key)
        try:
            ast = specialize_group(node.skeleton, bindings)
            result = self._evaluate_ast(ast)
            result = self._auto_resolve_single_cell(result)
            if result is None:
                result = 0
            self._cache[member_key] = result
            if self.on_cell_evaluated is not None:
                self.on_cell_evaluated(member_key, result)
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
        if isinstance(node, WholeColumnNode):
            return self._resolve_whole_column(node.sheet, node.column)
        if isinstance(node, WholeRowNode):
            return self._resolve_whole_row(node.sheet, node.row)
        if isinstance(node, RangeNode):
            return _range_from_a1(node.start, node.end)
        if isinstance(node, FunctionCallNode):
            name = normalize_excel_function_name(node.name)
            if name == "IF":
                return self._eval_if(node.args)
            if name == "IFERROR":
                return self._eval_iferror(node.args)
            if name == "IFNA":
                return self._eval_ifna(node.args)
            if name == "ISERROR":
                return self._eval_iserror(node.args)
            if name == "ISNA":
                return self._eval_isna(node.args)
            if name == "ISBLANK":
                return self._eval_isblank(node.args)
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
            if name == "INDEX":
                return self._eval_index(node.args)
            if name == "TRUE":
                return True
            if name == "FALSE":
                return False

            args = [self._evaluate_ast(a) for a in node.args]
            args = [self._resolve_function_arg(arg, name, index) for index, arg in enumerate(args)]
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
        return resolve_excel_range(rng, self._evaluate_cell)

    def _as_lazy_range(self, rng: ExcelRange) -> Range:
        """Bind an ``ExcelRange`` geometry to the evaluator cell resolver."""
        return Range(
            rng.sheet,
            rng.start_row,
            rng.start_col,
            rng.end_row,
            rng.end_col,
            self._evaluate_cell,
        )

    def _resolve_function_arg(
        self,
        value: CellValue,
        func_name: str,
        arg_index: int,
    ) -> CellValue:
        """Resolve ``ExcelRange`` arguments for runtime function calls.

        Policy (multi-cell):

        - ``eager_materialize_arg_indices`` → dense ndarray (full-scan bridge)
        - ``grid_range_arg_indices`` → lazy ``Range`` (selective Grid access)
        - otherwise → ``#VALUE!`` (scalar / non-Grid consumers; no Range leak)

        Single-cell references in value contexts (e.g. ``TEXT(INDEX(...))``)
        promote to scalars so export parity matches codegen's scalar ``INDEX``
        handling.
        """
        if not isinstance(value, ExcelRange):
            return value
        if arg_index in eager_materialize_arg_indices(func_name):
            return self._resolve_range(value)
        if value.start_row == value.end_row and value.start_col == value.end_col:
            return self._auto_resolve_single_cell(value)
        if arg_index in grid_range_arg_indices(func_name):
            # Lazy Range is a legal function operand; omitted from CellValue to
            # avoid a circular import with core.grid.
            return cast(CellValue, self._as_lazy_range(value))
        return XlError.VALUE

    def _sheet_bounds(self) -> SheetBounds:
        bounds = getattr(self.graph, "sheet_bounds", None)
        return dict(bounds) if bounds else {}

    def _resolve_whole_column(self, sheet: str, column: str) -> ExcelRange:
        return resolve_whole_column(sheet, column, self._sheet_bounds())

    def _resolve_whole_row(self, sheet: str, row: int) -> ExcelRange:
        return resolve_whole_row(sheet, row, self._sheet_bounds())

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

    def _resolve_binary_operand(self, value: CellValue) -> CellValue:
        """Resolve range references to arrays for element-wise binary operators."""
        if isinstance(value, ExcelRange):
            return self._resolve_range(value)
        return value

    def _eval_binary_op(self, node: BinaryOpNode) -> CellValue:
        left = self._resolve_binary_operand(self._evaluate_ast(node.left))
        right = self._resolve_binary_operand(self._evaluate_ast(node.right))

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

        # Arithmetic operators - element-wise when operands include arrays
        if op == "+":
            return xl_add(left, right)
        if op == "-":
            return xl_sub(left, right)
        if op == "*":
            return xl_mul(left, right)
        if op == "/":
            return xl_div(left, right)
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

    def _eval_isblank(self, args: list[AstNode]) -> bool:
        if len(args) != 1:
            raise ParseError("ISBLANK(...)", "ISBLANK requires 1 argument")
        v = self._evaluate_ast(args[0])
        if isinstance(v, ExcelRange):
            if v.start_row != v.end_row or v.start_col != v.end_col:
                return False
            v = self._resolve_range(v)[0, 0]
        return xl_isblank(v)

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

    def _eval_index(self, args: list[AstNode]) -> CellValue:
        node = FunctionCallNode(name="INDEX", args=args)
        ref = self._index_call_to_range(node)
        if isinstance(ref, XlError):
            return ref
        return ref

    def _index_call_to_range(self, node: FunctionCallNode) -> ExcelRange | XlError:
        if len(node.args) < 1:
            return XlError.VALUE
        array_node = node.args[0]
        if isinstance(array_node, WholeColumnNode):
            base = self._resolve_whole_column(array_node.sheet, array_node.column)
        elif isinstance(array_node, WholeRowNode):
            base = self._resolve_whole_row(array_node.sheet, array_node.row)
        elif isinstance(array_node, RangeNode):
            base = _range_from_a1(array_node.start, array_node.end)
        else:
            inner = self._range_from_ref_node(array_node)
            if not isinstance(inner, ExcelRange):
                return XlError.VALUE
            base = inner

        if len(node.args) < 2 or isinstance(node.args[1], EmptyArgNode):
            row_num = None
        else:
            row_num = self._evaluate_ast(node.args[1])
            if isinstance(row_num, XlError):
                return row_num
        if len(node.args) < 3 or isinstance(node.args[2], EmptyArgNode):
            col_num = None
        else:
            col_num = self._evaluate_ast(node.args[2])
            if isinstance(col_num, XlError):
                return col_num
        return index_excel_range(base, row_num, col_num)

    def _range_from_ref_node(self, node: AstNode) -> ExcelRange | XlError:
        """Interpret an AST node as a reference (cell or range) without evaluating its value."""
        if isinstance(node, RangeNode):
            return _range_from_a1(node.start, node.end)
        if isinstance(node, WholeColumnNode):
            return self._resolve_whole_column(node.sheet, node.column)
        if isinstance(node, WholeRowNode):
            return self._resolve_whole_row(node.sheet, node.row)

        if isinstance(node, CellRefNode):
            sheet, coord = parse_address(node.address)
            coord = coord.replace("$", "")
            col_str, row = fastpyxl.utils.cell.coordinate_from_string(coord)
            col = fastpyxl.utils.cell.column_index_from_string(col_str)
            return ExcelRange(sheet=sheet, start_row=row, start_col=col, end_row=row, end_col=col)

        if isinstance(node, FunctionCallNode) and node.name.upper() == "INDEX":
            return self._index_call_to_range(node)

        evaluated = self._evaluate_ast(node)
        if isinstance(evaluated, XlError):
            return evaluated
        if isinstance(evaluated, ExcelRange):
            return evaluated
        return XlError.VALUE


def _range_from_a1(start: str, end: str) -> ExcelRange:
    start_sheet, start_coord = parse_address(start)
    if "!" in end:
        end_sheet, end_coord = parse_address(end)
    else:
        end_sheet, end_coord = start_sheet, end
    if start_sheet != end_sheet:
        raise ValueError("Cross-sheet ranges are not supported")

    start_coord = start_coord.replace("$", "")
    end_coord = end_coord.replace("$", "")
    c1, r1 = fastpyxl.utils.cell.coordinate_from_string(start_coord)
    c2, r2 = fastpyxl.utils.cell.coordinate_from_string(end_coord)
    start_col = fastpyxl.utils.cell.column_index_from_string(c1)
    end_col = fastpyxl.utils.cell.column_index_from_string(c2)
    sr, er = sorted((r1, r2))
    sc, ec = sorted((start_col, end_col))
    return ExcelRange(sheet=start_sheet, start_row=sr, start_col=sc, end_row=er, end_col=ec)
