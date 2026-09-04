from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, cast, overload

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    format_key,
    parse_address,
)
from excel_grapher.core.address_keys import (
    normalize_key as normalize_address,
)
from excel_grapher.core.addressing import (
    index_excel_range,
    indirect_text_to_range,
    split_sheet_qualified_address,
)
from excel_grapher.core.excel_function_meta import grid_range_arg_indices
from excel_grapher.core.formula_ast import (
    bind_axes,
    resolve_cell_ref,
    resolve_whole_column_ref,
    resolve_whole_row_ref,
)
from excel_grapher.core.grid import Range
from excel_grapher.core.range_shorthand import (
    SheetBounds,
    resolve_whole_column,
    resolve_whole_row,
)
from excel_grapher.core.types import CellValue, ExcelRange, FormulaValue, XlError
from excel_grapher.evaluator.name_utils import normalize_excel_function_name
from excel_grapher.grapher.blank_ranges import (
    address_in_blank_ranges,
    normalize_blank_range_specs,
)
from excel_grapher.runtime.cache import (
    EvalContext,
    warn_circular_reference,
    xl_circular_reference,
    xl_iterative_compute,
)
from excel_grapher.runtime.info import xl_isblank
from excel_grapher.runtime.lookup import xl_index

from .ast_cache import DEFAULT_AST_CACHE_MAXSIZE, AstCache, AstCacheInfo
from .errors import MissingNormalizedFormulaError, ParseError
from .functions import FUNCTIONS
from .helpers import (
    get_error,
    to_bool,
    to_number,
    to_string,
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
    xl_neg,
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
from .shape_eval import ShapeEvalFn, compile_formula_shape

_SKIP_ERROR_PRECHECK = {
    # Selective Grid consumers: AST precheck must not force a full scan.
    "LOOKUP",
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "XLOOKUP",
    # Full-scan reductions that fail-fast (or Excel-skip) internally.
    "SUM",
    "AVERAGE",
    "MIN",
    "MAX",
    "STDEV",
    "NPV",
    "SUMPRODUCT",
    "LARGE",
    "RANK",
    "COUNT",
    "COUNTA",
    # Criteria consumers: Excel skips error cells in the criteria range.
    "COUNTIF",
    "AVERAGEIF",
    # Logical reductions: short-circuit and fail-fast over lazy ranges.
    "AND",
    "OR",
    # IS-family type predicates: Excel passes error values through; the
    # callables return FALSE (they are not numbers/text). ISERROR/ISNA/ISBLANK
    # are AST-special-cased above and never reach this generic path.
    "ISNUMBER",
    "ISTEXT",
}

if TYPE_CHECKING:
    from excel_grapher.grapher import DependencyGraph


@dataclass(frozen=True)
class _IndirectSheetBounds:
    """Workbook bounds for `indirect_text_to_range`, keyed to the target sheet."""

    sheet: str
    min_row: int = 1
    max_row: int = 1_048_576
    min_col: int = 1
    max_col: int = 16_384


@dataclass
class FormulaEvaluator:
    """Evaluate Excel formulas over a `DependencyGraph`.

    Holds a shared reference to `graph` (no deep copy). Cell-value and
    parsed-AST caches are per-evaluator.

    Formula-shape helpers (`_shape_fns`) are compiled in `__post_init__`
    from `graph.formula_shapes` at construction. That compile is a snapshot:
    reassigning or rewarming `graph.formula_shapes` later does not refresh
    `_shape_fns`. Construct a new `FormulaEvaluator` after rewarming if you
    want compiled shape helpers. Missing shapes fall back to
    `Node.formula_ast`; correctness does not require the overlay to be warm.
    """

    graph: DependencyGraph
    auto_detect_changes: bool = True
    eager_invalidation: bool = True
    on_cell_evaluated: Callable[[str, FormulaValue], None] | None = None
    iterate_enabled: bool = False
    iterate_count: int = 100
    iterate_delta: float = 0.001
    blank_ranges: tuple[str, ...] | None = None
    ast_cache_maxsize: int = DEFAULT_AST_CACHE_MAXSIZE

    def __post_init__(self) -> None:
        self._cache: dict[str, FormulaValue] = {}
        self._ast_cache = AstCache(maxsize=self.ast_cache_maxsize)
        self._seed_ast_cache_from_graph()
        self._shape_fns: dict[str, ShapeEvalFn] = {}
        table = getattr(self.graph, "formula_shapes", None)
        if table is not None:
            for shape_key, skeleton in table.shapes.items():
                self._shape_fns[shape_key] = compile_formula_shape(self, skeleton)
        self._call_stack: list[str] = []
        self._leaf_values: dict[str, FormulaValue] = {}  # For auto-detection
        self._iteration_values: dict[str, FormulaValue] = {}
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

    def _seed_ast_cache_from_graph(self) -> None:
        """Seed the string-keyed AST LRU from per-node trees and overlays.

        Per-node `formula_ast` is evaluated directly; this cache is the fallback
        parse path for formula cells that still only have `normalized_formula`.

        Cache contract: keys are stripped `normalized_formula` (absolute A1).
        That string is lossy for axis-aware trees, so values are always
        `bind_axes` copies resolved against the host `NodeKey`. Do not store a
        relative tree under this key. `preparsed_formulas` must use the same
        absolute-bound contract. `move_node` that preserves resolved targets
        can leave these entries valid; expire or rebuild when targets change.
        """
        entries: dict[str, AstNode] = {}
        for key, node in self.graph.formula_nodes():
            ast = node.formula_ast
            nf = node.normalized_formula
            if ast is None or not isinstance(nf, str):
                continue
            stripped = nf.strip()
            if stripped and stripped not in entries:
                host = node.address if node.address is not None else key
                entries[stripped] = bind_axes(ast, host)
        if entries:
            self._ast_cache.seed(entries)
        preparsed = getattr(self.graph, "preparsed_formulas", None)
        if preparsed:
            self._ast_cache.seed(preparsed)

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

    def _iterative_target_handler(self, addr: str) -> Callable[[EvalContext, str], FormulaValue]:
        def handler(_ctx: EvalContext, _target: str) -> FormulaValue:
            return self._evaluate_cell(addr)

        return handler

    @overload
    def evaluate(self, targets: str) -> FormulaValue: ...

    @overload
    def evaluate(self, targets: Sequence[str]) -> dict[str, FormulaValue]: ...

    def evaluate(self, targets: str | Sequence[str]) -> FormulaValue | dict[str, FormulaValue]:
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
            target_handlers: dict[str, Callable[[EvalContext, str], FormulaValue]] = {
                addr: self._iterative_target_handler(addr) for addr in target_list
            }
            ctx = EvalContext(
                inputs={},
                resolver=lambda _addr: None,
                cache=cast(dict[str, CellValue], self._cache),
                computing=set(self._call_stack),
                iterative_enabled=True,
                iterate_count=self.iterate_count,
                iterate_delta=self.iterate_delta,
                iteration_values=cast(dict[str, CellValue], self._iteration_values),
            )
            result = xl_iterative_compute(
                ctx,
                cast(
                    dict[str, Callable[[EvalContext, str], CellValue]],
                    target_handlers,
                ),
            )
            self._iteration_values = cast(dict[str, FormulaValue], ctx.iteration_values)
            if single:
                return cast(FormulaValue, next(iter(result.values())))
            return cast(dict[str, FormulaValue], result)
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

            if node.formula_ast is None and node.normalized_formula is None:
                # A raw formula without AST or normalized text is malformed, not a leaf.
                if node.formula is not None:
                    raise MissingNormalizedFormulaError(current)
                leaves.add(current)
            else:
                # Add its dependencies to the queue
                for dep in self._runtime_deps.get(current, set()):
                    if dep not in visited:
                        queue.append(dep)

        return leaves

    def _evaluate_cell(self, address: str) -> FormulaValue:
        norm = normalize_address(address)
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

        node = self.graph.get_node(norm)
        if node is None:
            raise KeyError(f"Cell {address} not found in graph")

        nf = node.normalized_formula
        formula_ast = getattr(node, "formula_ast", None)
        if formula_ast is None and nf is None:
            # A raw formula without AST or normalized text is malformed, not a leaf.
            if node.formula is not None:
                raise MissingNormalizedFormulaError(norm)
            self._cache[norm] = node.value
            self._leaf_values[norm] = node.value  # Track for change detection
            if self.on_cell_evaluated is not None:
                self.on_cell_evaluated(norm, node.value)
            return node.value

        formula: str | None = None
        if formula_ast is None:
            if not isinstance(nf, str) or not nf.strip():
                raise MissingNormalizedFormulaError(norm)
            formula = nf.strip()

        self._call_stack.append(norm)
        try:
            result = self._evaluate_formula(formula, node_key=norm, formula_ast=formula_ast)
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

    def _evaluate_formula(
        self,
        formula: str | None,
        *,
        node_key: str,
        formula_ast: AstNode | None = None,
    ) -> FormulaValue:
        """Evaluate a formula AST, preferring a compiled shape when interned.

        Shape helpers come from the construction-time `_shape_fns` snapshot.
        Missing overlay, missing binding, or a shape not compiled at init
        falls back to `formula_ast` / the string parse path.
        """
        table = getattr(self.graph, "formula_shapes", None)
        if table is not None:
            found = table.lookup(node_key)
            if found is not None:
                shape_key, _skeleton, params = found
                compiled = self._shape_fns.get(shape_key)
                if compiled is not None:
                    return compiled(params)
        if formula_ast is not None:
            return self._evaluate_ast(formula_ast)
        if formula is None:
            raise MissingNormalizedFormulaError(node_key)
        ast = self._parse_cached(formula)
        return self._evaluate_ast(ast)

    def _evaluate_ast(self, node: AstNode) -> FormulaValue:
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
            return self._evaluate_cell(resolve_cell_ref(node.ref, self._formula_anchor()))
        if isinstance(node, WholeColumnNode):
            sheet, column = resolve_whole_column_ref(node, self._formula_anchor())
            return self._resolve_whole_column(sheet, column)
        if isinstance(node, WholeRowNode):
            sheet, row = resolve_whole_row_ref(node, self._formula_anchor())
            return self._resolve_whole_row(sheet, row)
        if isinstance(node, RangeNode):
            start = resolve_cell_ref(node.start_ref, self._formula_anchor())
            end = resolve_cell_ref(node.end_ref, self._formula_anchor())
            return _range_from_a1(start, end)
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
            if name == "INDIRECT":
                return self._eval_indirect(node.args)
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
        value: FormulaValue,
        func_name: str,
        arg_index: int,
    ) -> FormulaValue:
        """Resolve ``ExcelRange`` arguments for runtime function calls.

        Policy (multi-cell):

        - ``grid_range_arg_indices`` → lazy ``Range`` (lookups + full-scan
          reductions); consumers walk cells via ``flatten`` / ``Grid``
        - otherwise → ``#VALUE!`` (scalar / non-Grid consumers; no Range leak)

        Single-cell references in value contexts (e.g. ``TEXT(INDEX(...))``)
        promote to scalars so export parity matches codegen's scalar ``INDEX``
        handling.
        """
        if not isinstance(value, ExcelRange):
            return value
        if value.start_row == value.end_row and value.start_col == value.end_col:
            return self._auto_resolve_single_cell(value)
        if arg_index in grid_range_arg_indices(func_name):
            # Lazy Range is a legal function operand; omitted from CellValue to
            # avoid a circular import with core.grid.
            return cast(FormulaValue, self._as_lazy_range(value))
        return XlError.VALUE

    def _single_cell_address(self, rng: ExcelRange) -> str:
        col = fastpyxl.utils.cell.get_column_letter(rng.start_col)
        return format_key(rng.sheet, f"{col}{rng.start_row}")

    def _auto_resolve_single_cell(self, value: FormulaValue) -> FormulaValue:
        """If value is a 1x1 ExcelRange, resolve it to its single cell value."""
        if (
            isinstance(value, ExcelRange)
            and value.start_row == value.end_row
            and value.start_col == value.end_col
        ):
            return self._evaluate_cell(self._single_cell_address(value))
        return value

    def _sheet_bounds(self) -> SheetBounds:
        bounds = getattr(self.graph, "sheet_bounds", None)
        return dict(bounds) if bounds else {}

    def _resolve_whole_column(self, sheet: str, column: str) -> ExcelRange:
        return resolve_whole_column(sheet, column, self._sheet_bounds())

    def _resolve_whole_row(self, sheet: str, row: int) -> ExcelRange:
        return resolve_whole_row(sheet, row, self._sheet_bounds())

    def _resolve_binary_operand(self, value: FormulaValue) -> FormulaValue:
        """Bind range geometry for element-wise operators.

        Single-cell references (including 1x1 results from `INDEX`) resolve to
        their scalar value so comparisons and concatenation match Excel scalar
        context. Multi-cell ranges stay as lazy `Range` for broadcast ops.
        """
        if isinstance(value, ExcelRange):
            if value.start_row == value.end_row and value.start_col == value.end_col:
                return self._auto_resolve_single_cell(value)
            return cast(FormulaValue, self._as_lazy_range(value))
        return value

    def _eval_binary_op(self, node: BinaryOpNode) -> FormulaValue:
        left = self._resolve_binary_operand(self._evaluate_ast(node.left))
        right = self._resolve_binary_operand(self._evaluate_ast(node.right))
        return self._apply_binary_op(node.op, left, right)

    def _apply_binary_op(self, op: str, left: FormulaValue, right: FormulaValue) -> FormulaValue:
        # Propagate errors
        if isinstance(left, XlError):
            return left
        if isinstance(right, XlError):
            return right

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

    def _eval_unary_op(self, node: UnaryOpNode) -> FormulaValue:
        operand = self._resolve_binary_operand(self._evaluate_ast(node.operand))
        return self._apply_unary_op(node.op, operand)

    def _apply_unary_op(self, op: str, operand: FormulaValue) -> FormulaValue:
        if isinstance(operand, XlError):
            return operand

        if op == "-":
            return xl_neg(operand)

        if op == "%":
            return xl_percent(operand)

        raise ValueError(f"Unknown unary operator: {op}")

    def _eval_if_branch(self, arg: AstNode) -> FormulaValue:
        """Evaluate an IF then/else branch.

        An empty argument (`IF(cond, a,)` / `IF(cond, , b)`) is Excel blank,
        which IF materializes as `0`. A truly omitted else (`IF(cond, a)`) is
        handled by the caller and yields `FALSE`.
        """
        if isinstance(arg, EmptyArgNode):
            return 0
        return self._evaluate_ast(arg)

    def _eval_if(self, args: Sequence[AstNode]) -> FormulaValue:
        if len(args) < 2:
            raise ParseError("IF(...)", "IF requires at least 2 arguments")
        cond = self._evaluate_ast(args[0])
        b = to_bool(cond)
        if isinstance(b, XlError):
            return b
        if b:
            return self._eval_if_branch(args[1])
        if len(args) >= 3:
            return self._eval_if_branch(args[2])
        return False

    def _eval_iferror(self, args: Sequence[AstNode]) -> FormulaValue:
        if len(args) < 2:
            raise ParseError("IFERROR(...)", "IFERROR requires 2 arguments")
        v = self._evaluate_ast(args[0])
        if isinstance(v, XlError):
            return self._evaluate_ast(args[1])
        return v

    def _eval_ifna(self, args: Sequence[AstNode]) -> FormulaValue:
        if len(args) < 2:
            raise ParseError("IFNA(...)", "IFNA requires 2 arguments")
        v = self._evaluate_ast(args[0])
        if v == XlError.NA:
            return self._evaluate_ast(args[1])
        return v

    def _eval_iserror(self, args: Sequence[AstNode]) -> bool:
        if len(args) < 1:
            raise ParseError("ISERROR(...)", "ISERROR requires 1 argument")
        v = self._evaluate_ast(args[0])
        return isinstance(v, XlError)

    def _eval_isna(self, args: Sequence[AstNode]) -> bool:
        if len(args) < 1:
            raise ParseError("ISNA(...)", "ISNA requires 1 argument")
        v = self._evaluate_ast(args[0])
        return v == XlError.NA

    def _eval_isblank(self, args: Sequence[AstNode]) -> bool:
        if len(args) != 1:
            raise ParseError("ISBLANK(...)", "ISBLANK requires 1 argument")
        v = self._evaluate_ast(args[0])
        if isinstance(v, ExcelRange):
            if v.start_row != v.end_row or v.start_col != v.end_col:
                return False
            v = self._evaluate_cell(self._single_cell_address(v))
        return xl_isblank(cast(CellValue, v))

    def _eval_choose(self, args: Sequence[AstNode]) -> FormulaValue:
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

    def _eval_offset(self, args: Sequence[AstNode]) -> FormulaValue:
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

        return xl_offset_ref(
            base,
            cast(CellValue, rows_val),
            cast(CellValue, cols_val),
            cast(CellValue, height_val) if height_val is not None else None,
            cast(CellValue, width_val) if width_val is not None else None,
        )

    def _eval_indirect(self, args: Sequence[AstNode]) -> FormulaValue:
        if len(args) < 1 or isinstance(args[0], EmptyArgNode):
            return XlError.VALUE
        text_val = self._evaluate_ast(args[0])
        if isinstance(text_val, XlError):
            return text_val
        text = to_string(text_val)
        a1 = True
        if len(args) >= 2 and not isinstance(args[1], EmptyArgNode):
            a1_val = self._evaluate_ast(args[1])
            if isinstance(a1_val, XlError):
                return a1_val
            flag = to_bool(a1_val)
            if isinstance(flag, XlError):
                return flag
            a1 = flag
        parsed = split_sheet_qualified_address(text.strip())
        if parsed is not None:
            sheet = parsed[0]
        else:
            anchor = self._formula_anchor()
            sheet = parse_address(anchor)[0] if anchor is not None else "Sheet1"
        bounds = _IndirectSheetBounds(sheet=sheet)
        return indirect_text_to_range(text, a1, bounds=bounds)

    def _formula_anchor(self) -> str | None:
        return self._call_stack[-1] if self._call_stack else None

    def _current_formula_row_col(self) -> tuple[int, int] | None:
        if not self._call_stack:
            return None
        _sheet, cell = parse_address(self._call_stack[-1])
        cell = cell.replace("$", "")
        col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
        col = fastpyxl.utils.cell.column_index_from_string(col_str)
        return row, col

    def _eval_row(self, args: Sequence[AstNode]) -> int | XlError:
        if not args or (len(args) == 1 and isinstance(args[0], EmptyArgNode)):
            pos = self._current_formula_row_col()
            return XlError.VALUE if pos is None else pos[0]
        ref = self._range_from_ref_node(args[0])
        if isinstance(ref, XlError):
            return ref
        return xl_row(ref)

    def _eval_column(self, args: Sequence[AstNode]) -> int | XlError:
        if not args or (len(args) == 1 and isinstance(args[0], EmptyArgNode)):
            pos = self._current_formula_row_col()
            return XlError.VALUE if pos is None else pos[1]
        ref = self._range_from_ref_node(args[0])
        if isinstance(ref, XlError):
            return ref
        return xl_column(ref)

    def _eval_columns(self, args: Sequence[AstNode]) -> int | XlError:
        if len(args) < 1:
            raise ParseError("COLUMNS(...)", "COLUMNS requires 1 argument")
        ref = self._range_from_ref_node(args[0])
        if isinstance(ref, XlError):
            return ref
        return xl_columns(ref)

    def _eval_index(self, args: Sequence[AstNode]) -> FormulaValue:
        """Evaluate INDEX for reference bases and computed value arrays.

        Literal ranges / nested INDEX/OFFSET keep the reference path
        (`ExcelRange` geometry). Computed arrays (e.g. `A1:A3<>0`) use shared
        `xl_index` / `index_cells` so evaluator and export agree (#503).
        """
        if len(args) < 1:
            return XlError.VALUE
        array_node = args[0]
        row_num, col_num, arg_err = self._index_row_col_args(args)
        if arg_err is not None:
            return arg_err

        if self._index_array_is_reference(array_node):
            base = self._index_base_range(array_node)
            if isinstance(base, ExcelRange):
                return index_excel_range(base, row_num, col_num)
            if isinstance(base, XlError) and base is not XlError.VALUE:
                return base
            # Nested INDEX may yield a computed array rather than geometry; fall
            # through to the shared value-INDEX path.

        array = self._evaluate_ast(array_node)
        if isinstance(array, XlError):
            return array
        if isinstance(array, ExcelRange):
            return index_excel_range(array, row_num, col_num)
        return cast(
            FormulaValue,
            xl_index(array, cast(CellValue, row_num), cast(CellValue, col_num)),
        )

    @staticmethod
    def _index_array_is_reference(node: AstNode) -> bool:
        """True when INDEX's array arg is a reference expression, not a value array."""
        if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode, CellRefNode)):
            return True
        if isinstance(node, FunctionCallNode):
            return node.name.upper() in {"INDEX", "OFFSET", "INDIRECT"}
        return False

    def _index_row_col_args(
        self, args: Sequence[AstNode]
    ) -> tuple[FormulaValue | None, FormulaValue | None, XlError | None]:
        if len(args) < 2 or isinstance(args[1], EmptyArgNode):
            row_num: FormulaValue | None = None
        else:
            row_num = self._evaluate_ast(args[1])
            if isinstance(row_num, XlError):
                return None, None, row_num
        if len(args) < 3 or isinstance(args[2], EmptyArgNode):
            col_num: FormulaValue | None = None
        else:
            col_num = self._evaluate_ast(args[2])
            if isinstance(col_num, XlError):
                return None, None, col_num
        return row_num, col_num, None

    def _index_base_range(self, array_node: AstNode) -> ExcelRange | XlError:
        """Resolve INDEX's array argument to `ExcelRange` geometry when possible."""
        if isinstance(array_node, WholeColumnNode):
            sheet, column = resolve_whole_column_ref(array_node, self._formula_anchor())
            return self._resolve_whole_column(sheet, column)
        if isinstance(array_node, WholeRowNode):
            sheet, row = resolve_whole_row_ref(array_node, self._formula_anchor())
            return self._resolve_whole_row(sheet, row)
        if isinstance(array_node, RangeNode):
            start = resolve_cell_ref(array_node.start_ref, self._formula_anchor())
            end = resolve_cell_ref(array_node.end_ref, self._formula_anchor())
            return _range_from_a1(start, end)
        return self._range_from_ref_node(array_node)

    def _index_call_to_range(self, node: FunctionCallNode) -> ExcelRange | XlError:
        if len(node.args) < 1:
            return XlError.VALUE
        base = self._index_base_range(node.args[0])
        if isinstance(base, XlError):
            return base
        row_num, col_num, arg_err = self._index_row_col_args(node.args)
        if arg_err is not None:
            return arg_err
        return index_excel_range(base, row_num, col_num)

    def _range_from_ref_node(self, node: AstNode) -> ExcelRange | XlError:
        """Interpret an AST node as a reference (cell or range) without evaluating its value."""
        if isinstance(node, RangeNode):
            start = resolve_cell_ref(node.start_ref, self._formula_anchor())
            end = resolve_cell_ref(node.end_ref, self._formula_anchor())
            return _range_from_a1(start, end)
        if isinstance(node, WholeColumnNode):
            sheet, column = resolve_whole_column_ref(node, self._formula_anchor())
            return self._resolve_whole_column(sheet, column)
        if isinstance(node, WholeRowNode):
            sheet, row = resolve_whole_row_ref(node, self._formula_anchor())
            return self._resolve_whole_row(sheet, row)

        if isinstance(node, CellRefNode):
            address = resolve_cell_ref(node.ref, self._formula_anchor())
            sheet, coord = parse_address(address)
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
