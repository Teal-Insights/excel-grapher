"""Code generator for converting Excel formulas to Python code."""

from __future__ import annotations

import ast as py_ast
import re
from collections.abc import Iterable, Mapping, Sequence, Set
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any, Protocol, TypedDict, cast

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    format_range_key,
    parse_address,
    parse_cell_coords,
    quote_sheet_if_needed,
    sort_node_keys,
)
from excel_grapher.core.address_keys import (
    normalize_key as normalize_address,
)
from excel_grapher.core.formula_ast import bind_axes, resolve_cell_ref
from excel_grapher.core.formula_shape import (
    AddressHoleNode,
    AddressLeaf,
    SkeletonNode,
    iter_address_holes,
    resolve_address_leaf,
    specialize_formula_shape,
)
from excel_grapher.core.operator_thresholds import MIN_OPERATOR_FASTPATH_CELLS
from excel_grapher.evaluator.errors import MissingNormalizedFormulaError
from excel_grapher.evaluator.name_utils import (
    address_to_python_name,
    excel_func_to_python,
    normalize_excel_function_name,
)
from excel_grapher.evaluator.parser import (
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
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.embed import emit_runtime
from excel_grapher.grapher.blank_ranges import BlankRangeRect, normalize_blank_range_specs
from excel_grapher.grapher.formula_label import display_formula
from excel_grapher.grapher.graph import CycleError
from excel_grapher.grapher.parser import format_key
from excel_grapher.grapher.target_expansion import (
    expand_targets_to_roots,
    split_range_target_on_colon,
)

__all__ = ["CodeGenerator", "GraphLike", "GraphNode"]

if TYPE_CHECKING:
    from excel_grapher.exporter.projection import ProjectionManifest
    from excel_grapher.grapher import DependencyGraph
    from excel_grapher.series_bindings.docstring_renderers import SeriesDocstringRendererSpec
    from excel_grapher.series_bindings.docstrings import SeriesBindingDocstringCallbackSpec
    from excel_grapher.series_bindings.output_helper_index import (
        OutputHelperIndex,
        OutputHelperSpec,
    )
    from excel_grapher.series_bindings.reader_index import ReaderIndex
    from excel_grapher.series_bindings.types import InputSeries, WorkbookSeriesBindings


class GraphNode(Protocol):
    formula: str | None
    normalized_formula: str | None
    formula_ast: AstNode | None
    value: object | None


_MAPPING_PROXY_IMPORT = "from types import MappingProxyType"


def _node_has_formula(node: object) -> bool:
    """Return True when `node` is a formula cell (AST or unparseable text)."""
    has = getattr(node, "has_formula", None)
    if isinstance(has, bool):
        return has
    return (
        getattr(node, "formula_ast", None) is not None
        or getattr(node, "normalized_formula", None) is not None
    )


def _ensure_mapping_proxy_import(source: str) -> str:
    """Insert `from types import MappingProxyType` after the `__future__` import."""
    if _MAPPING_PROXY_IMPORT in source:
        return source
    marker = "from __future__ import annotations"
    idx = source.find(marker)
    if idx == -1:
        return f"{_MAPPING_PROXY_IMPORT}\n\n{source}"
    insert_at = idx + len(marker)
    return f"{source[:insert_at]}\n\n{_MAPPING_PROXY_IMPORT}{source[insert_at:]}"


class GraphLike(Protocol):
    def get_node(self, address: str) -> GraphNode | None: ...

    def leaf_keys(self) -> list[str]: ...

    def formula_keys(self) -> list[str]: ...

    def get_dependencies(self, address: str) -> frozenset[str]: ...

    leaf_classification: dict[str, str] | None


class GenerationParts(TypedDict):
    runtime_code: str
    inputs_block_lines: list[str]
    constants_block_lines: list[str]
    cell_code_lines: list[str]
    formula_cells: list[str]
    all_cells: list[str]
    needs_offset_table: bool
    targets: list[str]
    has_constants: bool
    used_xl_functions: Set[str]
    blank_rects: tuple[BlankRangeRect, ...]


# Comparison operators emitted via xl_compare / xl_map_compare.
_COMPARE_OPS = frozenset({"=", "<>", "<", ">", "<=", ">="})

# Binary operators emitted as native Python with coercion helpers.
_ARITHMETIC_OPS = frozenset({"+", "-", "*", "/", "^"})

# Functions whose single argument is emitted as a lazily-evaluated thunk so the
# exported runtime can catch raised Excel errors. Mirrors the evaluator's
# AST-level special cases; other IS functions propagate argument errors there.
_THUNK_ARG_FUNCTIONS = frozenset({"ISERROR", "ISNA", "ISBLANK", "ISNUMBER", "ISTEXT"})

# Emit paths for these functions inspect concrete CellRefNode / RangeNode ASTs
# (static OFFSET, xl_index_ref, ROW of the formula cell). Shared helpers walk
# punched AddressHoleNode skeletons, so these stay on the filled-AST body.
_SHAPE_HELPER_REF_FUNCS = frozenset({"OFFSET", "INDEX", "ROW", "COLUMN", "COLUMNS"})

# Return unpacking hoists substantive ``xl_*`` runtime calls into statement-level
# temporaries. Coercion helpers and error literals stay inline because they are
# cheap and wrapping them would add noise without aiding debugging.
_RETURN_UNPACK_NON_HOISTABLE = frozenset(
    {
        "xl_number",
        "xl_bool",
        "xl_int",
        "xl_raise",
        "to_string",
    }
)


@dataclass
class _ReturnUnpackState:
    statements: list[str]


@dataclass
class _ReturnUnpackFrame:
    lazy: bool = False
    nested: bool = False


class CodeGenerator:
    """Generates Python code from Excel formulas.

    Reads `graph.formula_shapes` at `generate` time when present (not an
    init snapshot). Missing shapes fall back to per-node AST.
    """

    def __init__(
        self,
        graph: DependencyGraph | GraphLike,
        *,
        iterate_enabled: bool | None = None,
        iterate_count: int = 100,
        iterate_delta: float = 0.001,
        unpack_return: bool = False,
    ) -> None:
        """Initialize the code generator.

        Args:
            graph: Dependency graph from excel_grapher containing cell formulas.
            iterate_enabled: If True, `DependencyGraph.evaluation_order` rejects
                any must- or may-cycle (workbook iterative calc is unsupported in codegen).
                Typically set from `excel_grapher.get_calc_settings`. `None` skips
                this check (default).
            iterate_count: Maximum iterations when iterative calculation is enabled.
            iterate_delta: Convergence threshold when iterative calculation is enabled.
            unpack_return: When True, hoist nested runtime calls in each formula
                cell's return expression into statement-level temporaries.
        """
        self.graph = graph
        self._iterate_enabled = iterate_enabled
        self._unpack_return = unpack_return
        self._iterate_count = iterate_count
        self._iterate_delta = iterate_delta
        self._emitted: set[str] = set()
        self._needs_offset_runtime = False  # Set to True if dynamic OFFSET is used
        self._needs_index_ref_runtime = False  # OFFSET(INDEX(...), ...) requires xl_index_ref
        self._needs_operators_fastpath = False  # Large array binary ops / SUMPRODUCT
        self._offset_runtime_sheets: set[str] = set()
        self._temp_var_counter = 0  # Counter for unique temp variable names
        self._ast_cache: dict[str, AstNode] = {}
        self._used_graph_closure: bool = False
        self._formula_cell_address: str | None = None
        self._return_unpack_state: _ReturnUnpackState | None = None
        self._return_unpack_stack: list[_ReturnUnpackFrame] = []
        self._reader_index: ReaderIndex | None = None
        self._used_readers: set[str] = set()
        self._shape_helper_names: dict[str, str] = {}

    def __enter__(self) -> CodeGenerator:
        return self

    def __exit__(self, *args: object) -> None:
        self._reset_transient_state()
        return None

    def _reset_transient_state(self) -> None:
        self._emitted.clear()
        self._needs_offset_runtime = False
        self._needs_index_ref_runtime = False
        self._needs_operators_fastpath = False
        self._offset_runtime_sheets.clear()
        self._temp_var_counter = 0
        self._ast_cache.clear()
        self._used_graph_closure = False
        self._formula_cell_address = None
        self._return_unpack_state = None
        self._return_unpack_stack = []
        self._reader_index = None
        self._used_readers.clear()
        self._shape_helper_names.clear()

    def _include_dep_tracking(
        self,
        series_bindings: WorkbookSeriesBindings | None,
    ) -> bool:
        """Return whether exported runtime should embed dependency invalidation."""
        if self._iterate_enabled:
            return True
        if series_bindings is not None:
            from excel_grapher.series_bindings.normalize import has_input_direction

            for series in series_bindings.get("series", []):
                if isinstance(series, dict) and has_input_direction(series):
                    return True
        return False

    def _public_graph(self) -> DependencyGraph | GraphLike:
        original = getattr(self.graph, "original_graph", None)
        if original is not None:
            return cast("DependencyGraph | GraphLike", original)
        return self.graph

    def _projection_manifest(self) -> ProjectionManifest | None:
        return getattr(self.graph, "manifest", None)

    def _map_address_to_projected(self, address: str) -> str:
        manifest = self._projection_manifest()
        normalized = normalize_address(address)
        if manifest is None:
            return normalized
        return manifest.map_to_projected(normalized)

    def _projection_alias_map(
        self,
        public_addresses: Iterable[str],
        export_addresses: Iterable[str],
    ) -> dict[str, str]:
        manifest = self._projection_manifest()
        if manifest is None:
            return {}
        exported = frozenset(normalize_address(addr) for addr in export_addresses)
        aliases: dict[str, str] = {}
        for address in public_addresses:
            public_addr = normalize_address(address)
            projected_addr = normalize_address(manifest.map_to_projected(public_addr))
            if projected_addr != public_addr and projected_addr in exported:
                aliases[public_addr] = projected_addr
        return aliases

    def _export_addresses_with_aliases(
        self,
        export_addresses: Iterable[str],
        public_addresses: Iterable[str],
    ) -> list[str]:
        addresses = [normalize_address(addr) for addr in export_addresses]
        alias_map = self._projection_alias_map(public_addresses, addresses)
        if not alias_map:
            return addresses
        merged = list(addresses)
        seen = set(addresses)
        for alias in sorted(alias_map):
            if alias not in seen:
                merged.append(alias)
                seen.add(alias)
        return merged

    def _series_binding_public_addresses(
        self,
        bindings: WorkbookSeriesBindings | None,
        workbook: Path | str | None,
    ) -> frozenset[str]:
        if bindings is None:
            return frozenset()
        if workbook is None:
            raise ValueError("bindings_workbook is required when series_bindings is set")
        from excel_grapher.series_bindings.workflow import series_binding_public_addresses

        return series_binding_public_addresses(
            cast("DependencyGraph", self._public_graph()),
            bindings,
            workbook=workbook,
        )

    def _should_emit_compute_all(
        self,
        targets: Sequence[str],
        *,
        series_bindings: WorkbookSeriesBindings | None,
        bindings_workbook: Path | str | None,
        export_addresses: Iterable[str] | None,
        include_compute_all: bool | None,
    ) -> bool:
        """Return whether public `compute_all` / `TARGETS` should be emitted."""
        from excel_grapher.series_bindings.workflow import (
            output_binding_covered_addresses,
            should_emit_compute_all,
        )

        covered: frozenset[str] = frozenset()
        if series_bindings is not None:
            if bindings_workbook is None:
                raise ValueError("bindings_workbook is required when series_bindings is set")
            covered = output_binding_covered_addresses(
                cast("DependencyGraph", self._public_graph()),
                series_bindings,
                workbook=bindings_workbook,
                export_addresses=export_addresses,
            )
        return should_emit_compute_all(
            targets,
            covered_by_output=covered,
            include_compute_all=include_compute_all,
        )

    def _emit_compute_all_block(self, targets: Sequence[str]) -> list[str]:
        """Emit `TARGETS` map and public `compute_all` entry point."""
        lines = [
            "TARGETS = {",
            *[
                f"    {repr(target)}: {handler},"
                for target, handler in self._targets_to_entries(targets)
            ],
            "}",
            "",
            "",
            (
                "def compute_all(ctx: EvalContext | None = None, *, "
                "inputs: dict[str, object] | None = None) -> dict[str, object]:"
            ),
            '    """Compute all target cells and return results."""',
            "    if ctx is None:",
            "        ctx = make_context(inputs)",
            "    elif inputs is not None:",
            (
                "        warnings.warn("
                '"inputs will be ignored because ctx was provided", '
                "UserWarning, stacklevel=2)"
            ),
        ]
        if self._iterate_enabled:
            lines.append("    return xl_iterative_compute(ctx, TARGETS)")
        else:
            lines.append(
                "    return {target: handler(ctx, target) for target, handler in TARGETS.items()}"
            )
        lines.append("")
        return lines

    def _emit_projection_alias_lines(
        self,
        export_addresses: Iterable[str],
        public_addresses: Iterable[str],
    ) -> list[str]:
        alias_map = self._projection_alias_map(public_addresses, export_addresses)
        if not alias_map:
            return []
        lines = ["# --- Projection public address aliases ---", ""]
        for public_addr, replacement in sorted(alias_map.items()):
            public_fn = address_to_python_name(public_addr)
            lines.append(f"def {public_fn}(ctx):")
            lines.extend(self._emit_projection_alias_body(replacement))
            lines.append("")
        return lines

    def _emit_projection_alias_body(self, replacement: str) -> list[str]:
        """Emit the body lines for a projected public-address alias wrapper.

        Always delegate to the retained target's exported evaluation path so the
        alias shares memoization with other callers instead of duplicating its
        formula body.
        """
        unpack_stmts = self._start_return_unpack()
        try:
            expr = self._emit_cell_eval(replacement)
        finally:
            self._stop_return_unpack()
        return self._format_return_lines(unpack_stmts, expr)

    def _graph_sheetnames(self, *, targets: Sequence[str] | None = None) -> list[str]:
        sheet_order = getattr(self.graph, "sheet_order", None)
        if sheet_order:
            return list(sheet_order)

        sheets: list[str] = []
        seen: set[str] = set()

        def _add_sheet(sheet: str) -> None:
            if sheet not in seen:
                seen.add(sheet)
                sheets.append(sheet)

        def _add_from_keys(keys: Sequence[str]) -> None:
            for key in keys:
                try:
                    sheet, _ = parse_address(normalize_address(key))
                except ValueError:
                    continue
                _add_sheet(sheet)

        keys_fn = getattr(self.graph, "keys", None)
        if callable(keys_fn):
            _add_from_keys(list(keys_fn()))
        else:
            for attr in ("leaf_keys", "formula_keys"):
                fn = getattr(self.graph, attr, None)
                if callable(fn):
                    _add_from_keys(list(fn()))

        if targets:
            for raw in targets:
                token = str(raw)
                if "!" not in token:
                    continue
                split = split_range_target_on_colon(token)
                start = split[0] if split is not None else token
                try:
                    sheet, _ = parse_address(start)
                except ValueError:
                    continue
                _add_sheet(sheet)

        return sheets

    def _named_range_maps(
        self,
    ) -> tuple[dict[str, tuple[str, str]], dict[str, tuple[str, str, str]]]:
        public_graph = self._public_graph()
        named_ranges = getattr(public_graph, "named_ranges", None) or {}
        named_range_ranges = getattr(public_graph, "named_range_ranges", None) or {}
        return named_ranges, named_range_ranges

    def _expand_target_tokens(self, targets: Sequence[str]) -> list[str]:
        named_ranges, named_range_ranges = self._named_range_maps()
        roots = expand_targets_to_roots(
            targets,
            sheetnames=self._graph_sheetnames(targets=targets),
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
        return [normalize_address(format_key(sheet, a1)) for sheet, a1 in roots]

    def _get_or_parse_ast(self, address: str) -> AstNode | None:
        """Parse and cache the AST for a formula cell.

        The cache key is the normalized, sheet-qualified cell address. The cache
        is cleared at the start of each generate() call.
        """
        normalized = normalize_address(address)
        if normalized in self._ast_cache:
            return self._ast_cache[normalized]

        node = self.graph.get_node(normalized)
        if node is None:
            return None

        nf = node.normalized_formula
        formula_ast = node.formula_ast
        if formula_ast is None and nf is None:
            # A raw formula without AST or normalized text is malformed, not a leaf.
            if node.formula is not None:
                raise MissingNormalizedFormulaError(normalized)
            return None
        if formula_ast is not None:
            bound = bind_axes(formula_ast, normalized)
            self._ast_cache[normalized] = bound
            return bound
        if not isinstance(nf, str) or not nf.strip():
            raise MissingNormalizedFormulaError(normalized)
        table = getattr(self.graph, "formula_shapes", None)
        if table is not None:
            found = table.lookup(normalized)
            if found is not None:
                _shape_key, skeleton, params = found
                ast = bind_axes(specialize_formula_shape(skeleton, params), normalized)
                self._ast_cache[normalized] = ast
                return ast
        ast = parse(nf.strip())
        self._ast_cache[normalized] = ast
        return ast

    def _emit_ast(self, node: SkeletonNode) -> str:
        """Convert an AST node to a Python expression string.

        Args:
            node: AST node to convert.

        Returns:
            Python expression as a string.
        """
        if isinstance(node, EmptyArgNode):
            return "None"

        if isinstance(node, NumberNode):
            return repr(node.value)

        if isinstance(node, StringNode):
            return repr(node.value)

        if isinstance(node, BoolNode):
            return "True" if node.value else "False"

        if isinstance(node, ErrorNode):
            # Error literals raise in the exported error channel.
            return f"xl_raise(XlError.{node.error.name})"

        if isinstance(node, AddressHoleNode):
            return self._emit_address_hole(node)

        if isinstance(node, CellRefNode):
            return self._emit_cell_eval(node.address)

        if isinstance(node, RangeNode):
            return self._emit_range(node)

        if isinstance(node, BinaryOpNode):
            return self._emit_binary_op(node)

        if isinstance(node, UnaryOpNode):
            return self._emit_unary_op(node)

        if isinstance(node, FunctionCallNode):
            return self._emit_function_call(node)

        raise ValueError(f"Unknown AST node type: {type(node)}")

    def _emit_range(self, node: RangeNode) -> str:
        """Emit a range as a lazy `Range` value resolved through the context.

        For A1:B3, emits: xl_range(ctx, "S!A1:B3"). Consumers evaluate cells
        positionally; unused cells are never evaluated.

        A 1x1 range collapses to a scalar cell read so binary/unary operators
        match Excel and the evaluator (issue #421).
        """
        if self._range_node_is_single_cell(node):
            return self._emit_cell_eval(node.start)
        return self._emit_range_address(node.start, node.end)

    def _emit_range_address(self, start: str, end: str) -> str:
        """Emit an xl_range or binding-aligned read_*_range call for a start/end pair."""
        sheet, r1, c1, r2, c2 = self._range_coords(start, end)
        start_cell = f"{fastpyxl.utils.cell.get_column_letter(c1)}{r1}"
        end_cell = f"{fastpyxl.utils.cell.get_column_letter(c2)}{r2}"
        range_key = format_range_key(sheet, start_cell, end_cell)
        if self._reader_index is not None:
            from excel_grapher.series_bindings.reader_index import resolve_reader_ref

            resolved = resolve_reader_ref(range_key, index=self._reader_index)
            if resolved["reader"] is not None:
                self._used_readers.add(resolved["reader"])
            expr = resolved["call_form"]
        else:
            expr = f"xl_range(ctx, {repr(range_key)})"
        return self._hoist_return_expr(expr)

    def _emit_cell_eval(self, address: str) -> str:
        normalized = normalize_address(address)
        if self.graph is None:
            expr = f"xl_cell(ctx, {repr(normalized)})"
        else:
            node = self.graph.get_node(normalized)
            if node is not None and _node_has_formula(node):
                func_name = address_to_python_name(normalized)
                expr = f"xl_eval(ctx, {repr(normalized)}, {func_name})"
            elif self._reader_index is not None:
                from excel_grapher.series_bindings.reader_index import resolve_reader_ref

                resolved = resolve_reader_ref(normalized, index=self._reader_index)
                if resolved["reader"] is not None:
                    self._used_readers.add(resolved["reader"])
                expr = resolved["call_form"]
            else:
                expr = f"xl_cell(ctx, {repr(normalized)})"
        return self._hoist_return_expr(expr)

    def _emit_address_hole(self, node: AddressHoleNode) -> str:
        """Emit a punched address hole as `xl_cell` / `xl_range` on a helper param.

        The bound address is a runtime string, so helpers cannot emit
        `xl_eval(ctx, addr, cell_fn)`. `xl_cell` resolves formula cells through
        `ctx.resolver` and leaf cells through inputs.
        """
        pname = f"p{node.index}"
        expr = f"xl_cell(ctx, {pname})" if node.kind == "CELL" else f"xl_range(ctx, {pname})"
        return self._hoist_return_expr(expr)

    @staticmethod
    def _py_literal(value: Any) -> str:
        """Convert a Python value into a safe Python literal expression.

        The generated code must be syntactically valid Python. Values pulled from
        workbooks can include objects (e.g., fastpyxl ArrayFormula) whose repr()
        is not a literal and would break the generated file if embedded.
        """
        if value is None:
            return "0"
        if isinstance(value, XlError):
            return f"XlError.{value.name}"
        if isinstance(value, (bool, int, float, str)):
            return repr(value)
        # Numpy scalars may appear; keep the runtime surface small by emitting their
        # native Python equivalent when available.
        if hasattr(value, "item"):
            try:
                return CodeGenerator._py_literal(value.item())
            except Exception:
                return "0"
        return "0"

    def _emit_series_binding_setters(
        self,
        bindings: WorkbookSeriesBindings,
        workbook: Path | str,
        *,
        export_addresses: Iterable[str],
        public_addresses: Iterable[str],
        include_helpers: bool = True,
        include_readers: bool = True,
        include_leaf_indexes: bool = True,
        include_leaves_tables: bool = True,
        series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
        docstring_renderer: SeriesDocstringRendererSpec = "google",
        helper_index: OutputHelperIndex | None = None,
        address_helpers: Mapping[str, OutputHelperSpec] | None = None,
    ) -> list[str]:
        from excel_grapher.series_bindings.bindings_codegen import emit_series_bindings_block

        return emit_series_bindings_block(
            cast("DependencyGraph", self._public_graph()),
            workbook,
            bindings,
            export_addresses=self._export_addresses_with_aliases(
                export_addresses,
                public_addresses,
            ),
            include_helpers=include_helpers,
            include_readers=include_readers,
            include_leaf_indexes=include_leaf_indexes,
            include_leaves_tables=include_leaves_tables,
            series_docstring_callback=series_docstring_callback,
            docstring_renderer=docstring_renderer,
            helper_index=helper_index,
            address_helpers=address_helpers,
        )

    @staticmethod
    def _series_bindings_have_input(bindings: WorkbookSeriesBindings) -> bool:
        """Return True when any series declares an input (setter) direction."""
        from excel_grapher.series_bindings.normalize import has_input_direction

        return any(
            isinstance(series, dict) and has_input_direction(series)
            for series in bindings.get("series", [])
        )

    @staticmethod
    def _series_bindings_have_readers(bindings: WorkbookSeriesBindings) -> bool:
        """Return True when any series emits a public `read_*` (input or constant)."""
        from excel_grapher.series_bindings.normalize import has_reader_direction

        return any(
            isinstance(series, dict) and has_reader_direction(series)
            for series in bindings.get("series", [])
        )

    @staticmethod
    def _series_bindings_have_output(bindings: WorkbookSeriesBindings) -> bool:
        """Return True when any series declares an output (compute) direction."""
        from excel_grapher.series_bindings.normalize import has_output_direction

        return any(
            isinstance(series, dict) and has_output_direction(series)
            for series in bindings.get("series", [])
        )

    @staticmethod
    def _series_bindings_may_emit_range_readers(bindings: WorkbookSeriesBindings) -> bool:
        """Return True when reader series may emit `read_*_range` helpers needing `xl_range`."""
        from excel_grapher.series_bindings.normalize import has_reader_direction

        for series in bindings.get("series", []):
            if not isinstance(series, dict) or not has_reader_direction(series):
                continue
            if series.get("layout") == "scalar":
                continue
            from excel_grapher.series_bindings.ranges import series_data_ranges

            for data_range in series_data_ranges(series):
                if ":" in data_range:
                    return True
        return False

    @staticmethod
    def _series_binding_emitted_range_reader_names(lines: Sequence[str]) -> list[str]:
        """Extract `read_*_range` function names from emitted bindings code."""
        names: list[str] = []
        for line in lines:
            match = re.match(r"^def (read_[a-z0-9_]+_range)\(", line)
            if match:
                name = match.group(1)
                if name not in names:
                    names.append(name)
        return names

    @staticmethod
    def _emit_api_helpers_module() -> str:
        """Emit the `_api_helpers.py` module holding series-binding coercion helpers."""
        from excel_grapher.series_bindings.setter_codegen import (
            SERIES_HELPERS_STDLIB_IMPORTS,
            emit_series_helpers_definitions,
        )

        lines: list[str] = [
            "from __future__ import annotations",
            "",
            *SERIES_HELPERS_STDLIB_IMPORTS,
            "",
            "from .runtime import coerce_inputs_dict",
            "",
            *emit_series_helpers_definitions(),
        ]
        return "\n".join(lines).rstrip() + "\n"

    @staticmethod
    def _emit_output_leaves_module(leaf_lines: Sequence[str]) -> str:
        """Emit the `_output_leaves.py` module holding `_OUTPUT_LEAVES_*` tables."""
        lines: list[str] = [
            "from __future__ import annotations",
            "",
            *leaf_lines,
        ]
        return "\n".join(lines).rstrip() + "\n"

    @staticmethod
    def _series_output_leaves_imports(lines: Sequence[str]) -> list[str]:
        """Return `_OUTPUT_LEAVES_*` symbols that `api.py` needs to import."""
        names: list[str] = []
        for line in lines:
            match = re.match(r"^(_OUTPUT_LEAVES_[A-Z0-9_]+):", line)
            if match:
                name = match.group(1)
                if name not in names:
                    names.append(name)
        return names

    @staticmethod
    def _emit_readers_module(reader_lines: Sequence[str]) -> str:
        """Emit the `_readers.py` module holding leaf maps and read_* duals."""
        text = "\n".join(reader_lines)
        runtime_names = ["CellValue", "EvalContext", "xl_cell"]
        if "xl_range(" in text:
            runtime_names.append("xl_range")
        runtime_import = CodeGenerator._format_from_runtime_import(runtime_names)
        lines: list[str] = [
            "from __future__ import annotations",
            "",
            runtime_import,
            "",
            *reader_lines,
        ]
        return "\n".join(lines).rstrip() + "\n"

    @staticmethod
    def _series_reader_leaf_index_imports(lines: Sequence[str]) -> list[str]:
        """Return `_LEAF_INDEX_*` symbols that setters in `api.py` need to import."""
        names: list[str] = []
        for line in lines:
            match = re.match(r"^(_LEAF_INDEX_[A-Z0-9_]+) =", line)
            if match:
                name = match.group(1)
                if name not in names:
                    names.append(name)
        return names

    @staticmethod
    def _series_reader_public_imports(lines: Sequence[str]) -> list[str]:
        """Return public `read_*` symbols defined in `_readers` for package re-export."""
        names: list[str] = []
        for line in lines:
            match = re.match(r"^def (read_[a-z0-9_]+)\(", line)
            if match:
                name = match.group(1)
                if name not in names:
                    names.append(name)
        return names

    @staticmethod
    def _format_from_module_import(
        module: str,
        names: list[str],
        *,
        noqa: str | None = None,
    ) -> str:
        """Format a relative `from .<module> import ...` statement."""
        if not names:
            return ""
        joined = ", ".join(names)
        prefix = f"from .{module} import "
        suffix = f"  # noqa: {noqa}" if noqa else ""
        if len(prefix) + len(joined) + len(suffix) <= 88:
            return prefix + joined + suffix
        inner = ",\n    ".join(names)
        if noqa:
            return f"{prefix}(  # noqa: {noqa}\n    {inner},\n)"
        return f"{prefix}(\n    {inner},\n)"

    _SERIES_HELPER_IMPORT_NAMES: tuple[str, ...] = (
        "DataFrameInput",
        "EmptyMeasure",
        "Record",
        "Records",
        "Scalar",
        "Sequence",
        "SeriesInput",
        "_apply_series_records",
        "coerce_setter_input",
    )

    @classmethod
    def _series_helper_imports(cls, lines: Sequence[str]) -> list[str]:
        """Return the helper names referenced by emitted setter/compute code."""
        text = "\n".join(lines)
        return [
            name
            for name in cls._SERIES_HELPER_IMPORT_NAMES
            if re.search(rf"\b{re.escape(name)}\b", text)
        ]

    def derive_input_series(
        self,
        bindings: WorkbookSeriesBindings,
        *,
        workbook: Path | str,
    ) -> list[InputSeries]:
        """Derive input-series metadata from explicit series bindings."""
        from excel_grapher.series_bindings import derive_input_series

        return derive_input_series(
            cast("DependencyGraph", self._public_graph()), bindings, workbook=workbook
        )

    @staticmethod
    def _series_binding_public_names(
        bindings: WorkbookSeriesBindings,
    ) -> tuple[list[str], list[str], list[str]]:
        """Return declared public setter, reader, and compute function names.

        Without groups the names sort alphabetically (flat export); with
        view-level groups they follow the grouped export order.
        """
        from excel_grapher.series_bindings.groups import (
            bindings_have_groups,
            grouped_public_names,
        )
        from excel_grapher.series_bindings.workflow import (
            compute_names,
            reader_names,
            setter_names,
        )

        if bindings_have_groups(bindings):
            return grouped_public_names(bindings)
        return setter_names(bindings), reader_names(bindings), compute_names(bindings)

    @staticmethod
    def _series_binding_groups_manifest(
        bindings: WorkbookSeriesBindings | None,
    ) -> dict[str, Any] | None:
        """Return the group manifest when any binding declares groups."""
        if bindings is None:
            return None
        from excel_grapher.series_bindings.groups import bindings_have_groups, group_manifest

        if not bindings_have_groups(bindings):
            return None
        return dict(group_manifest(bindings))

    @staticmethod
    def _series_binding_reader_discovery(
        graph: DependencyGraph,
        bindings: WorkbookSeriesBindings | None,
        *,
        workbook: Path | str | None,
        export_addresses: Iterable[str] | None = None,
    ) -> tuple[dict[str, dict[str, object]] | None, dict[str, dict[str, object]] | None]:
        """Return discovery payloads for `list_reader_leaves` / `list_reader_ranges`."""
        if bindings is None or workbook is None:
            return None, None
        from excel_grapher.series_bindings.normalize import has_reader_direction
        from excel_grapher.series_bindings.reader_index import (
            build_reader_index,
            reader_index_as_discovery_dicts,
        )

        if not any(
            isinstance(series, dict) and has_reader_direction(series)
            for series in bindings.get("series", [])
        ):
            return None, None
        index = build_reader_index(
            graph,
            bindings,
            workbook=workbook,
            export_addresses=export_addresses,
        )
        return reader_index_as_discovery_dicts(index)

    @staticmethod
    def _emit_series_binding_discovery_lines(
        setter_names: Sequence[str],
        compute_names: Sequence[str],
        groups_manifest: Mapping[str, Any] | None = None,
        reader_names: Sequence[str] | None = None,
        reader_leaves: Mapping[str, Mapping[str, object]] | None = None,
        reader_ranges: Mapping[str, Mapping[str, object]] | None = None,
    ) -> list[str]:
        """Emit generated-code helpers that list public series-binding functions."""
        lines = [
            "def list_setters() -> list[str]:",
            '    """Return generated series-binding setter function names."""',
            f"    return {list(setter_names)!r}",
            "",
            "",
            "def list_readers() -> list[str]:",
            '    """Return generated series-binding reader function names."""',
            f"    return {list(reader_names or ())!r}",
            "",
            "",
            "def list_computes() -> list[str]:",
            '    """Return generated series-binding compute function names."""',
            f"    return {list(compute_names)!r}",
        ]
        if reader_leaves is not None:
            lines.extend(
                [
                    "",
                    "",
                    "def list_reader_leaves() -> dict[str, dict[str, object]]:",
                    '    """Return address → semantic reader call metadata."""',
                    f"    return {dict(reader_leaves)!r}",
                ]
            )
        if reader_ranges is not None:
            lines.extend(
                [
                    "",
                    "",
                    "def list_reader_ranges() -> dict[str, dict[str, object]]:",
                    '    """Return binding-aligned data_range → range-reader metadata."""',
                    f"    return {dict(reader_ranges)!r}",
                ]
            )
        if groups_manifest is not None:
            lines.extend(
                [
                    "",
                    "",
                    "def list_groups() -> dict[str, object]:",
                    '    """Return the view-level group manifest for the generated API."""',
                    f"    return {dict(groups_manifest)!r}",
                ]
            )
        return lines

    def _range_addresses_2d(self, start: str, end: str) -> list[list[str]]:
        """Generate all cell addresses in a range as a 2D list (rows x cols)."""
        start_sheet, start_cell = self._parse_address(start)
        end_sheet, end_cell = self._parse_address(end)

        # Use start sheet for all cells (Excel semantics)
        sheet = start_sheet

        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell)

        start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)

        rows = []
        for row in range(start_row, end_row + 1):
            row_addrs = []
            for col_idx in range(start_col_idx, end_col_idx + 1):
                col_letter = fastpyxl.utils.cell.get_column_letter(col_idx)
                row_addrs.append(f"{sheet}!{col_letter}{row}")
            rows.append(row_addrs)

        return rows

    def _range_addresses(self, start: str, end: str) -> list[str]:
        """Generate all cell addresses in a range (flat list)."""
        rows = self._range_addresses_2d(start, end)
        return [addr for row in rows for addr in row]

    @staticmethod
    def _format_cell_address(sheet: str, row: int, col: int) -> str:
        sheet_name = quote_sheet_if_needed(sheet)
        col_letter = fastpyxl.utils.cell.get_column_letter(col)
        return f"{sheet_name}!{col_letter}{row}"

    def _targets_to_entries(self, targets: Sequence[str]) -> list[tuple[str, str]]:
        by_sheet: dict[str, list[tuple[int, int]]] = {}
        for address in targets:
            sheet, cell = parse_address(address)
            col_letters, row = fastpyxl.utils.cell.coordinate_from_string(cell)
            col_idx = fastpyxl.utils.cell.column_index_from_string(col_letters)
            by_sheet.setdefault(sheet, []).append((row, col_idx))

        entries: list[tuple[str, str]] = []

        for sheet, cells in by_sheet.items():
            cells_set = set(cells)
            if not cells_set:
                continue

            row_groups: dict[int, list[int]] = {}
            col_groups: dict[int, list[int]] = {}
            for row, col in cells_set:
                row_groups.setdefault(row, []).append(col)
                col_groups.setdefault(col, []).append(row)

            row_entries: list[tuple[str, str]] = []
            for row, cols in row_groups.items():
                cols = sorted(cols)
                start = prev = cols[0]
                for col in cols[1:]:
                    if col == prev + 1:
                        prev = col
                        continue
                    if start == prev:
                        row_entries.append(
                            (self._format_cell_address(sheet, row, start), "xl_cell")
                        )
                    else:
                        start_cell = f"{fastpyxl.utils.cell.get_column_letter(start)}{row}"
                        end_cell = f"{fastpyxl.utils.cell.get_column_letter(prev)}{row}"
                        row_entries.append(
                            (format_range_key(sheet, start_cell, end_cell), "xl_range_rows")
                        )
                    start = prev = col
                if start == prev:
                    row_entries.append((self._format_cell_address(sheet, row, start), "xl_cell"))
                else:
                    start_cell = f"{fastpyxl.utils.cell.get_column_letter(start)}{row}"
                    end_cell = f"{fastpyxl.utils.cell.get_column_letter(prev)}{row}"
                    row_entries.append(
                        (format_range_key(sheet, start_cell, end_cell), "xl_range_rows")
                    )

            col_entries: list[tuple[str, str]] = []
            for col, rows in col_groups.items():
                rows = sorted(rows)
                start = prev = rows[0]
                for row in rows[1:]:
                    if row == prev + 1:
                        prev = row
                        continue
                    if start == prev:
                        col_entries.append(
                            (self._format_cell_address(sheet, start, col), "xl_cell")
                        )
                    else:
                        col_letter = fastpyxl.utils.cell.get_column_letter(col)
                        col_entries.append(
                            (
                                format_range_key(
                                    sheet, f"{col_letter}{start}", f"{col_letter}{prev}"
                                ),
                                "xl_range_rows",
                            )
                        )
                    start = prev = row
                if start == prev:
                    col_entries.append((self._format_cell_address(sheet, start, col), "xl_cell"))
                else:
                    col_letter = fastpyxl.utils.cell.get_column_letter(col)
                    col_entries.append(
                        (
                            format_range_key(sheet, f"{col_letter}{start}", f"{col_letter}{prev}"),
                            "xl_range_rows",
                        )
                    )

            entries.extend(row_entries if len(row_entries) <= len(col_entries) else col_entries)

        entries.sort(key=lambda item: item[0])
        return entries

    @staticmethod
    def _emit_blank_range_lines(rects: tuple[BlankRangeRect, ...]) -> list[str]:
        """Emit compact structural-blank handling for declared blank rectangles."""
        rect_repr = ",\n    ".join(
            f"({repr(sh)}, {r1}, {c1}, {r2}, {c2})" for sh, r1, c1, r2, c2 in rects
        )
        return [
            "# --- Declared structural blank ranges ---",
            "_BLANK_RANGE_RECTS = (",
            f"    {rect_repr},",
            ")",
            "",
            "def _blank_range_parse_address(address):",
            '    if address.startswith("\'"):',
            "        i = 1",
            "        while i < len(address):",
            '            if address[i] == "\'":',
            '                if i + 1 < len(address) and address[i + 1] == "\'":',
            "                    i += 2",
            "                    continue",
            "                break",
            "            i += 1",
            '        sheet = address[1:i].replace("\'\'", "\'")',
            "        rest = address[i + 1 :]",
            '        if not rest.startswith("!"):',
            '            raise ValueError(f"Invalid address: {address!r}")',
            "        return sheet, rest[1:]",
            '    if "!" in address:',
            '        sheet, cell = address.rsplit("!", 1)',
            "        return sheet, cell",
            '    raise ValueError(f"Address must be sheet-qualified: {address!r}")',
            "",
            "def _address_in_blank_ranges(address):",
            "    sheet, cell = _blank_range_parse_address(address)",
            "    col_s, row = fastpyxl.utils.cell.coordinate_from_string(cell)",
            "    col = fastpyxl.utils.cell.column_index_from_string(col_s)",
            "    for sh, r1, c1, r2, c2 in _BLANK_RANGE_RECTS:",
            "        if sh != sheet:",
            "            continue",
            "        if r1 <= row <= r2 and c1 <= col <= c2:",
            "            return True",
            "    return False",
            "",
            "def _blank_structural_cell(ctx):",
            "    return None",
            "",
            "_blank_structural_cell.__structural_blank__ = True",
            "",
        ]

    @classmethod
    def _emit_resolver_lines(
        cls, blank_rects: tuple[BlankRangeRect, ...] | None = None
    ) -> list[str]:
        prefix: list[str] = []
        if blank_rects:
            prefix = cls._emit_blank_range_lines(blank_rects)
        resolve_head = [
            "# --- Formula resolver ---",
            "_RESOLVED_FORMULAS = {}",
            "def _address_to_func_name(address):",
            "    name = []",
            "    prev_underscore = False",
            "    for ch in address.lower():",
            '        if ch == "\'":',
            "            continue",
            '        if "a" <= ch <= "z" or "0" <= ch <= "9":',
            "            name.append(ch)",
            "            prev_underscore = False",
            "        else:",
            "            if not prev_underscore:",
            '                name.append("_")',
            "                prev_underscore = True",
            '    base = "".join(name).strip("_")',
            '    return f"cell_{base}"',
            "",
            "def _resolve_formula(address):",
        ]
        if blank_rects:
            resolve_head.extend(
                [
                    "    if _address_in_blank_ranges(address):",
                    "        return _blank_structural_cell",
                ]
            )
        resolve_head.extend(
            [
                "    fn = _RESOLVED_FORMULAS.get(address)",
                "    if fn is not None:",
                "        return fn",
                "    name = _address_to_func_name(address)",
                "    fn = globals().get(name)",
                "    if fn is not None:",
                "        _RESOLVED_FORMULAS[address] = fn",
                "    return fn",
                "",
            ]
        )
        return prefix + resolve_head

    @staticmethod
    def _internals_runtime_import_names(
        used_xl_functions: Set[str], cell_code_lines: list[str]
    ) -> list[str]:
        """Names from the embedded runtime that formula cell bodies reference as globals."""
        blob = "\n".join(cell_code_lines)
        if "def " not in blob:
            return []
        names = set(used_xl_functions)
        # Only pull symbols that appear in emitted bodies. After Phase 2, bound
        # leaves may use `read_*` exclusively, so `xl_cell` is not always needed.
        for symbol in ("xl_cell", "xl_eval", "xl_range", "xl_raise", "XlError", "ExcelRange"):
            if symbol == "XlError":
                if "XlError" in blob:
                    names.add(symbol)
            elif f"{symbol}(" in blob:
                names.add(symbol)
        return sorted(names)

    @staticmethod
    def _internals_needs_datetime_import(cell_code_lines: list[str]) -> bool:
        """Return True when formula bodies emit `datetime.datetime(...)` literals."""
        return "datetime.datetime" in "\n".join(cell_code_lines)

    @staticmethod
    def _format_from_runtime_import(names: list[str]) -> str:
        if not names:
            return ""
        joined = ", ".join(names)
        prefix = "from .runtime import "
        if len(prefix) + len(joined) <= 88:
            return prefix + joined
        inner = ",\n    ".join(names)
        return f"{prefix}(\n    {inner},\n)"

    def _parse_address(self, address: str) -> tuple[str, str]:
        """Parse a sheet-qualified address into (quoted_sheet, cell) tuple.

        The sheet name is returned with quotes if needed for address construction.
        """
        sheet, cell = parse_address(address)
        return quote_sheet_if_needed(sheet), cell

    def _get_graph_leaf_classification(self) -> dict[str, str] | None:
        mapping = getattr(self.graph, "leaf_classification", None)
        if mapping is None:
            return None
        if not isinstance(mapping, Mapping):
            raise TypeError("leaf_classification must be a mapping of address to label")
        normalized: dict[str, str] = {}
        for key, value in mapping.items():
            if not isinstance(key, str):
                raise TypeError("leaf_classification keys must be strings")
            if value not in {"input", "constant"}:
                raise ValueError("leaf_classification values must be 'input' or 'constant'")
            normalized[normalize_address(key)] = value
        return normalized

    def _collect_needed_leaves(self, all_cells: list[str]) -> set[str]:
        # Only emit leaf inputs that are actually needed for the target dependency closure.
        # This keeps generated output small and avoids embedding unrelated workbook artifacts.
        needed_leaves: set[str] = set()
        for addr in all_cells:
            node = self.graph.get_node(addr)
            if node is None or _node_has_formula(node):
                continue
            needed_leaves.add(normalize_address(addr))
        return needed_leaves

    @staticmethod
    def _normalize_constant_types(constant_types: set[str] | None) -> set[str]:
        if constant_types is None:
            return set()
        if isinstance(constant_types, (str, bytes)):
            raise TypeError("constant_types must be a set of strings")
        normalized = {str(item) for item in constant_types}
        allowed = {"number", "string"}
        invalid = normalized - allowed
        if invalid:
            raise ValueError(f"Unsupported constant_types: {sorted(invalid)!r}")
        return normalized

    @classmethod
    def _classification_from_graph(
        cls, graph_classification: dict[str, str] | None, needed_leaves: set[str]
    ) -> tuple[set[str], set[str]]:
        if graph_classification is None:
            return set(needed_leaves), set()
        constants = {addr for addr in needed_leaves if graph_classification.get(addr) == "constant"}
        inputs = set(needed_leaves) - constants
        return inputs, constants

    @staticmethod
    def _parse_constant_range(range_str: str) -> tuple[str, int, int, int, int]:
        if not isinstance(range_str, str):
            raise TypeError("constant_ranges entries must be strings")
        if "!" not in range_str:
            raise ValueError(f"Range must be sheet-qualified: {range_str}")
        sheet_part, cell_part = range_str.rsplit("!", 1)
        if ":" in cell_part:
            start_cell, end_cell = cell_part.split(":", 1)
        else:
            start_cell = end_cell = cell_part

        sheet, start = parse_address(f"{sheet_part}!{start_cell}")
        _, end = parse_address(f"{sheet_part}!{end_cell}")

        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start)
        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end)
        start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)

        r1, r2 = (start_row, end_row) if start_row <= end_row else (end_row, start_row)
        c1, c2 = (
            (start_col_idx, end_col_idx)
            if start_col_idx <= end_col_idx
            else (end_col_idx, start_col_idx)
        )
        return (sheet, r1, c1, r2, c2)

    @classmethod
    def _normalize_constant_ranges(
        cls, constant_ranges: Sequence[str] | None
    ) -> list[tuple[str, int, int, int, int]]:
        if constant_ranges is None:
            return []
        if isinstance(constant_ranges, (str, bytes)):
            raise TypeError("constant_ranges must be a sequence of strings")
        return [cls._parse_constant_range(item) for item in constant_ranges]

    @classmethod
    def _normalize_input_ranges(
        cls, input_ranges: Sequence[str] | None
    ) -> list[tuple[str, int, int, int, int]]:
        if input_ranges is None:
            return []
        if isinstance(input_ranges, (str, bytes)):
            raise TypeError("input_ranges must be a sequence of strings")
        return [cls._parse_constant_range(item) for item in input_ranges]

    @staticmethod
    def _apply_input_ranges_override(
        needed_leaves: set[str],
        constants: set[str],
        input_ranges: list[tuple[str, int, int, int, int]],
    ) -> tuple[set[str], set[str]]:
        """Drop constants that fall in input_ranges; input ranges win over constant rules."""
        if not input_ranges:
            return set(needed_leaves) - constants, constants
        constants = set(constants)
        for key in needed_leaves:
            if CodeGenerator._leaf_in_constant_ranges(key, input_ranges):
                constants.discard(key)
        inputs = set(needed_leaves) - constants
        return inputs, constants

    @staticmethod
    def _leaf_value_matches_constant_type(value: object | None, constant_types: set[str]) -> bool:
        if not constant_types:
            return False
        if value is None:
            value = 0
        if isinstance(value, bool):
            return False
        return ("number" in constant_types and isinstance(value, (int, float))) or (
            "string" in constant_types and isinstance(value, str)
        )

    @staticmethod
    def _leaf_in_constant_ranges(
        address: str, constant_ranges: list[tuple[str, int, int, int, int]]
    ) -> bool:
        if not constant_ranges:
            return False
        sheet, cell = parse_address(normalize_address(address))
        col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
        col = fastpyxl.utils.cell.column_index_from_string(col_str)
        for range_sheet, r1, c1, r2, c2 in constant_ranges:
            if sheet != range_sheet:
                continue
            if r1 <= row <= r2 and c1 <= col <= c2:
                return True
        return False

    def classify_leaf_nodes(
        self,
        targets: list[str],
        *,
        constant_types: set[str] | None = None,
        constant_ranges: Sequence[str] | None = None,
        constant_blanks: bool = False,
        input_ranges: Sequence[str] | None = None,
        attach_to_graph: bool = False,
    ) -> tuple[set[str], set[str]]:
        normalized_targets = [normalize_address(t) for t in targets]
        all_cells = self._collect_all_cells(normalized_targets)
        needed_leaves = self._collect_needed_leaves(all_cells)

        normalized_constant_types = self._normalize_constant_types(constant_types)
        normalized_constant_ranges = self._normalize_constant_ranges(constant_ranges)
        normalized_input_ranges = self._normalize_input_ranges(input_ranges)
        explicit_constant_rules = bool(constant_types or constant_ranges or constant_blanks)
        use_graph_classification = not explicit_constant_rules and not input_ranges

        if use_graph_classification:
            graph_classification = self._get_graph_leaf_classification()
            inputs, constants = self._classification_from_graph(graph_classification, needed_leaves)
        elif explicit_constant_rules:
            inputs, constants = self._classify_leaf_nodes(
                needed_leaves,
                constant_types=normalized_constant_types,
                constant_ranges=normalized_constant_ranges,
                constant_blanks=constant_blanks,
                input_ranges=normalized_input_ranges,
            )
        else:
            graph_classification = self._get_graph_leaf_classification()
            inputs, constants = self._classification_from_graph(graph_classification, needed_leaves)
            inputs, constants = self._apply_input_ranges_override(
                needed_leaves, constants, normalized_input_ranges
            )

        if attach_to_graph:
            classification = {addr: "input" for addr in inputs}
            classification.update({addr: "constant" for addr in constants})
            self.graph.leaf_classification = classification

        return inputs, constants

    def _classify_leaf_nodes(
        self,
        needed_leaves: set[str],
        *,
        constant_types: set[str],
        constant_ranges: list[tuple[str, int, int, int, int]],
        constant_blanks: bool,
        input_ranges: list[tuple[str, int, int, int, int]] | None = None,
    ) -> tuple[set[str], set[str]]:
        input_ranges = input_ranges or []
        constants: set[str] = set()
        for key in needed_leaves:
            if self._leaf_in_constant_ranges(key, constant_ranges):
                constants.add(key)
                continue
            node = self.graph.get_node(key)
            value = None if node is None else node.value
            if constant_blanks and value is None:
                constants.add(key)
                continue
            if self._leaf_value_matches_constant_type(value, constant_types):
                constants.add(key)
        return self._apply_input_ranges_override(needed_leaves, constants, input_ranges)

    def _binary_operator_exprs(self, op: str, left: str, right: str) -> tuple[str, str]:
        """Return `(scalar_expr, array_expr)` for a binary operator over two operands.

        The operand strings are substituted verbatim; callers that emit the
        array guard pass bound temp-variable names so each operand is rendered
        (and evaluated) once.
        """
        if op == "&":
            return f"(to_string({left}) + to_string({right}))", f"xl_map_concat({left}, {right})"
        if op in _ARITHMETIC_OPS:
            if op in {"+", "-", "*"}:
                scalar = f"(xl_number({left}) {op} xl_number({right}))"
            elif op == "/":
                scalar = (
                    f"((lambda _ln, _rn: (_ln / _rn if _rn != 0 else xl_raise(XlError.DIV)))"
                    f"(xl_number({left}), xl_number({right})))"
                )
            else:
                scalar = f"xl_pow_numbers(xl_number({left}), xl_number({right}))"
            return scalar, f"xl_map_arithmetic({op!r}, {left}, {right})"
        if op in _COMPARE_OPS:
            return (
                f"xl_compare({op!r}, {left}, {right})",
                f"xl_map_compare({op!r}, {left}, {right})",
            )
        raise ValueError(f"Unknown operator: {op}")

    def _emit_binary_op(self, node: BinaryOpNode) -> str:
        """Emit a binary operation with inlined scalar operators.

        When operands may be arrays at runtime, the operator branches between
        broadcast and scalar handling. Operands are bound to temp variables via
        a lambda so each is evaluated once, keeping generated code linear in the
        operator-tree depth instead of tripling per nesting level.
        """
        left = self._emit_ast_child(node.left)
        right = self._emit_ast_child(node.right)
        op = node.op

        if not self._ast_needs_array_operator_branch(node):
            scalar, _ = self._binary_operator_exprs(op, left, right)
            return scalar

        lname = self._next_temp_var()
        rname = self._next_temp_var()
        scalar, array = self._binary_operator_exprs(op, lname, rname)
        guard = f"({array} if (xl_is_array({lname}) or xl_is_array({rname})) else {scalar})"
        return f"(lambda {lname}, {rname}: {guard})({left}, {right})"

    def _unary_scalar_expr(self, op: str, operand: str) -> str:
        if op == "-":
            return f"(-xl_number({operand}))"
        if op == "+":
            return f"(+xl_number({operand}))"
        if op == "%":
            return f"(xl_number({operand}) / 100.0)"
        raise ValueError(f"Unknown unary operator: {op}")

    def _emit_unary_op(self, node: UnaryOpNode) -> str:
        """Emit a unary operation with inlined scalar operators.

        Like `_emit_binary_op`, the operand is bound once when an array branch
        is possible so the operand is not re-emitted across guard branches.
        """
        operand = self._emit_ast_child(node.operand)
        op = node.op

        if not self._ast_needs_array_operator_branch(node):
            return self._unary_scalar_expr(op, operand)

        name = self._next_temp_var()
        scalar = self._unary_scalar_expr(op, name)
        array = f"xl_map_unary({op!r}, {name})"
        guard = f"({array} if xl_is_array({name}) else {scalar})"
        return f"(lambda {name}: {guard})({operand})"

    def _emit_function_call(self, node: FunctionCallNode) -> str:
        """Emit a function call.

        Range arguments pass through as lazy `Range` values; IF, IFERROR/IFNA,
        CHOOSE, OFFSET, INDEX, ROW/COLUMN/COLUMNS are handled specially.
        """
        func_name = excel_func_to_python(node.name)
        upper_name = normalize_excel_function_name(node.name)

        # IF needs special handling - emit as Python conditional for lazy evaluation
        if upper_name == "IF":
            return self._emit_if(node)

        # IFERROR / IFNA need lazy fallback branches (not plain runtime calls).
        if upper_name in {"IFERROR", "IFNA"}:
            return self._emit_lazy_error_fallback(node, upper_name)

        # CHOOSE needs special handling - only evaluate the selected argument
        if upper_name == "CHOOSE":
            return self._emit_choose(node)

        # OFFSET needs special handling - try static resolution first
        if upper_name == "OFFSET":
            return self._emit_offset(node)

        # INDEX over a literal range: use reference indexing when the result is a single cell
        if (
            upper_name == "INDEX"
            and node.args
            and isinstance(node.args[0], RangeNode)
            and self._index_range_result_is_scalar(node.args[0], node)
        ):
            return self._emit_index_scalar_range(node)

        # ROW needs special handling - references should not be evaluated
        if upper_name == "ROW":
            return self._emit_row(node)

        # COLUMN needs special handling - references should not be evaluated
        if upper_name == "COLUMN":
            return self._emit_column(node)

        # COLUMNS needs special handling - references should not be evaluated
        if upper_name == "COLUMNS":
            return self._emit_columns(node)

        # TRUE()/FALSE() as zero-arg function calls
        if upper_name == "TRUE":
            return "True"
        if upper_name == "FALSE":
            return "False"

        # NA() is the functional spelling of the #N/A literal: raise in the
        # export error channel so the code never leaks as a sentinel into a
        # generic consumer (matching the evaluator's argument error precheck).
        if upper_name == "NA":
            return "xl_raise(XlError.NA)"

        # IS functions must not propagate errors: the argument is passed as a
        # lazily-evaluated thunk so the runtime can catch raised Excel errors.
        if upper_name in _THUNK_ARG_FUNCTIONS and len(node.args) == 1:
            with self._return_unpack_lazy():
                arg_expr = self._emit_ast_child(node.args[0])
            return f"{func_name}(lambda: ({arg_expr}))"

        emitted_args = [self._emit_ast_child(arg) for arg in node.args]
        args = ", ".join(emitted_args)
        expr = f"{func_name}({args})"
        if self._is_hoistable_runtime_func(func_name):
            return self._hoist_return_expr(expr)
        return expr

    def _next_temp_var(self) -> str:
        """Generate a unique temporary variable name."""
        self._temp_var_counter += 1
        return f"_t{self._temp_var_counter}"

    @staticmethod
    def _is_hoistable_runtime_func(name: str) -> bool:
        return name not in _RETURN_UNPACK_NON_HOISTABLE and (
            name.startswith("xl_") or name == "ExcelRange"
        )

    def _hoist_return_expr(self, expr: str, *, hoistable: bool = True) -> str:
        """Assign a nested runtime expression to a return-level temporary."""
        if (
            not self._unpack_return
            or self._return_unpack_state is None
            or not self._return_unpack_stack
            or self._return_unpack_stack[-1].lazy
            or not self._return_unpack_stack[-1].nested
            or not hoistable
        ):
            return expr
        name = self._next_temp_var()
        self._return_unpack_state.statements.append(f"{name} = {expr}")
        return name

    @contextmanager
    def _return_unpack_lazy(self):
        if not self._return_unpack_stack:
            yield
            return
        frame = self._return_unpack_stack[-1]
        prev_lazy = frame.lazy
        frame.lazy = True
        try:
            yield
        finally:
            frame.lazy = prev_lazy

    def _emit_ast_child(self, node: SkeletonNode) -> str:
        """Emit a nested formula operand while optionally unpacking return temps."""
        if not self._return_unpack_stack:
            return self._emit_ast(node)
        self._return_unpack_stack.append(
            _ReturnUnpackFrame(lazy=self._return_unpack_stack[-1].lazy, nested=True)
        )
        try:
            return self._emit_ast(node)
        finally:
            self._return_unpack_stack.pop()

    def _emit_lazy_error_fallback(self, node: FunctionCallNode, name: str) -> str:
        """Emit IFERROR/IFNA as thunked runtime calls with try/except semantics.

        IFERROR(value, value_if_error) evaluates the fallback only when
        evaluating ``value`` produces any Excel error (raised
        ``XlErrorException`` or ``XlError`` sentinel). IFNA does so only for
        ``#N/A`` and re-raises other errors.
        """
        if len(node.args) < 2:
            return "xl_raise(XlError.VALUE)"

        with self._return_unpack_lazy():
            value_expr = self._emit_ast_child(node.args[0])
            fallback_expr = self._emit_ast_child(node.args[1])

        if name == "IFERROR":
            func = "xl_iferror"
        elif name == "IFNA":
            func = "xl_ifna"
        else:
            raise ValueError(f"Unsupported lazy error fallback function: {name!r}")

        expr = f"{func}(lambda: ({value_expr}), lambda: ({fallback_expr}))"
        return self._hoist_return_expr(expr)

    def _emit_if(self, node: FunctionCallNode) -> str:
        """Emit IF as a Python conditional expression for lazy evaluation.

        IF(condition, true_val, [false_val])

        Emits as a nested conditional that:
        1. Returns error if condition is an error
        2. Otherwise lazily evaluates only the relevant branch

        This ensures only the relevant branch is evaluated, which is critical
        for breaking circular references that Excel handles via lazy evaluation.
        """
        if len(node.args) < 2:
            return "xl_raise(XlError.VALUE)"

        cond_expr = self._emit_ast_child(node.args[0])
        with self._return_unpack_lazy():
            # Empty IF branches (trailing/interior commas) are Excel blank -> 0.
            # A truly omitted else (`IF(cond, a)`) defaults to FALSE.
            true_expr = (
                "0"
                if isinstance(node.args[1], EmptyArgNode)
                else self._emit_ast_child(node.args[1])
            )
            if len(node.args) > 2:
                false_expr = (
                    "0"
                    if isinstance(node.args[2], EmptyArgNode)
                    else self._emit_ast_child(node.args[2])
                )
            else:
                false_expr = "False"

        # Excel-style boolean coercion is not Python truthiness:
        # - "FALSE" should behave like False
        # - "0" should produce #VALUE! (per to_bool)
        # `xl_bool` keeps lazy branch evaluation while raising coercion errors.
        bool_var = self._next_temp_var()
        return f"(({true_expr}) if ({bool_var} := xl_bool({cond_expr})) else ({false_expr}))"

    def _emit_row(self, node: FunctionCallNode) -> str:
        if not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)):
            addr = self._formula_cell_address
            if addr is None:
                return "xl_raise(XlError.VALUE)"
            _sheet, cell = parse_address(addr)
            cell_clean = cell.replace("$", "")
            _col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell_clean)
            return repr(int(row))

        arg = node.args[0]
        if isinstance(arg, CellRefNode):
            sheet, cell = parse_address(arg.address)
            col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
            col = fastpyxl.utils.cell.column_index_from_string(col_str)
            return f"xl_row(ExcelRange({repr(sheet)}, {row}, {col}, {row}, {col}))"
        if isinstance(arg, RangeNode):
            sheet, r1, c1, r2, c2 = self._range_coords(arg.start, arg.end)
            return f"xl_row(ExcelRange({repr(sheet)}, {r1}, {c1}, {r2}, {c2}))"
        if isinstance(arg, FunctionCallNode) and arg.name.upper() == "OFFSET":
            return f"xl_row({self._emit_offset_ref(arg)})"

        return f"xl_row({self._emit_ast_child(arg)})"

    def _emit_column(self, node: FunctionCallNode) -> str:
        if not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)):
            addr = self._formula_cell_address
            if addr is None:
                return "xl_raise(XlError.VALUE)"
            _sheet, cell = parse_address(addr)
            cell_clean = cell.replace("$", "")
            col_str, _row = fastpyxl.utils.cell.coordinate_from_string(cell_clean)
            col = fastpyxl.utils.cell.column_index_from_string(col_str)
            return repr(int(col))

        arg = node.args[0]
        if isinstance(arg, CellRefNode):
            sheet, cell = parse_address(arg.address)
            col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
            col = fastpyxl.utils.cell.column_index_from_string(col_str)
            return f"xl_column(ExcelRange({repr(sheet)}, {row}, {col}, {row}, {col}))"
        if isinstance(arg, RangeNode):
            sheet, r1, c1, r2, c2 = self._range_coords(arg.start, arg.end)
            return f"xl_column(ExcelRange({repr(sheet)}, {r1}, {c1}, {r2}, {c2}))"
        if isinstance(arg, FunctionCallNode) and arg.name.upper() == "OFFSET":
            return f"xl_column({self._emit_offset_ref(arg)})"

        return f"xl_column({self._emit_ast_child(arg)})"

    def _emit_columns(self, node: FunctionCallNode) -> str:
        if len(node.args) < 1:
            return "xl_raise(XlError.VALUE)"

        arg = node.args[0]
        if isinstance(arg, CellRefNode):
            sheet, cell = parse_address(arg.address)
            col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
            col = fastpyxl.utils.cell.column_index_from_string(col_str)
            return f"xl_columns(ExcelRange({repr(sheet)}, {row}, {col}, {row}, {col}))"
        if isinstance(arg, RangeNode):
            sheet, r1, c1, r2, c2 = self._range_coords(arg.start, arg.end)
            return f"xl_columns(ExcelRange({repr(sheet)}, {r1}, {c1}, {r2}, {c2}))"
        if isinstance(arg, FunctionCallNode) and arg.name.upper() == "OFFSET":
            return f"xl_columns({self._emit_offset_ref(arg)})"

        return f"xl_columns({self._emit_ast_child(arg)})"

    def _emit_choose(self, node: FunctionCallNode) -> str:
        """Emit CHOOSE as chained conditionals for lazy evaluation.

        CHOOSE(index_num, value1, [value2], ...)

        Emits as chained conditionals that only evaluate the selected value.
        This is critical for breaking circular references that Excel handles
        via lazy evaluation.
        """
        if len(node.args) < 2:
            return "xl_raise(XlError.VALUE)"

        index_expr = self._emit_ast_child(node.args[0])
        with self._return_unpack_lazy():
            value_exprs = [self._emit_ast_child(arg) for arg in node.args[1:]]

        # Store index in a temp var to avoid evaluating twice and to keep typing clean.
        # `xl_int` performs Excel-style numeric coercion and raises on errors.
        idx_var = self._next_temp_var()

        # Build chained conditionals: if idx==1 then val1 else if idx==2 then val2 ...
        # Start from the innermost (last value or a raised VALUE error when out of bounds)
        result = "xl_raise(XlError.VALUE)"
        for i, val_expr in reversed(list(enumerate(value_exprs, start=1))):
            result = f"(({val_expr}) if {idx_var} == {i} else ({result}))"

        # Wrap with bounds checking; coercion errors raise from `xl_int`.
        return (
            f"((xl_raise(XlError.VALUE) if ({idx_var} := xl_int({index_expr})) < 1 "
            f"or {idx_var} > {len(value_exprs)} else {result}))"
        )

    def _is_constant_number(self, node: AstNode) -> bool:
        """Check if a node is a constant numeric value.

        Handles both NumberNode and unary negation of NumberNode (e.g., -2).
        """
        if isinstance(node, NumberNode):
            return True
        if isinstance(node, UnaryOpNode) and node.op == "-":
            return isinstance(node.operand, NumberNode)
        return False

    def _get_constant_number(self, node: AstNode) -> float:
        """Extract the numeric value from a constant number node.

        Assumes _is_constant_number() has already returned True.
        """
        if isinstance(node, NumberNode):
            return node.value
        if (
            isinstance(node, UnaryOpNode)
            and node.op == "-"
            and isinstance(node.operand, NumberNode)
        ):
            return -node.operand.value
        raise ValueError(f"Not a constant number: {node}")

    def _can_offset_be_static(self, node: FunctionCallNode) -> bool:
        """Check if an OFFSET call can be statically resolved.

        Returns True if reference is a cell or range and all offsets/sizes are constants.
        """
        if len(node.args) < 3:
            return False

        ref_node = node.args[0]
        rows_node = node.args[1]
        cols_node = node.args[2]
        height_node = node.args[3] if len(node.args) > 3 else None
        width_node = node.args[4] if len(node.args) > 4 else None

        return (
            isinstance(ref_node, (CellRefNode, RangeNode))
            and self._is_constant_number(rows_node)
            and self._is_constant_number(cols_node)
            and (height_node is None or self._is_constant_number(height_node))
            and (width_node is None or self._is_constant_number(width_node))
        )

    def _static_offset_is_multicell(self, node: FunctionCallNode) -> bool:
        """True when a statically resolvable OFFSET produces a multi-cell range."""
        if not self._can_offset_be_static(node):
            return False
        ref_node = node.args[0]
        assert isinstance(ref_node, (CellRefNode, RangeNode))
        height_node = node.args[3] if len(node.args) > 3 else None
        width_node = node.args[4] if len(node.args) > 4 else None
        base_h, base_w = self._offset_base_shape(ref_node)
        height = int(self._get_constant_number(height_node)) if height_node is not None else base_h
        width = int(self._get_constant_number(width_node)) if width_node is not None else base_w
        return height != 1 or width != 1

    def _emit_offset(self, node: FunctionCallNode) -> str:
        """Emit OFFSET function, trying static resolution first.

        OFFSET(reference, rows, cols, [height], [width])

        If all offset arguments are constants, resolves to direct cell/range reference.
        Otherwise, falls back to runtime xl_offset() function.
        """
        if len(node.args) < 3:
            # Invalid OFFSET - need at least reference, rows, cols
            return "xl_raise(XlError.VALUE)"

        ref_node = node.args[0]
        rows_node = node.args[1]
        cols_node = node.args[2]
        height_node = node.args[3] if len(node.args) > 3 else None
        width_node = node.args[4] if len(node.args) > 4 else None

        # Try static resolution if reference is a cell and offsets are constants
        if self._can_offset_be_static(node):
            assert isinstance(ref_node, (CellRefNode, RangeNode))
            base_address = ref_node.address if isinstance(ref_node, CellRefNode) else ref_node.start
            base_h, base_w = self._offset_base_shape(ref_node)
            height = (
                int(self._get_constant_number(height_node)) if height_node is not None else base_h
            )
            width = int(self._get_constant_number(width_node)) if width_node is not None else base_w
            return self._emit_offset_static(
                base_address,
                int(self._get_constant_number(rows_node)),
                int(self._get_constant_number(cols_node)),
                height,
                width,
            )

        # Fall back to runtime resolution
        self._needs_offset_runtime = True
        return self._emit_offset_dynamic(ref_node, rows_node, cols_node, height_node, width_node)

    def _offset_base_shape(self, ref_node: AstNode) -> tuple[int, int]:
        """Return (height, width) for an OFFSET base reference."""
        if isinstance(ref_node, CellRefNode):
            return (1, 1)
        if isinstance(ref_node, RangeNode):
            _, r1, c1, r2, c2 = self._range_coords(ref_node.start, ref_node.end)
            return (r2 - r1 + 1, c2 - c1 + 1)
        return (1, 1)

    @staticmethod
    def _index_arg_is_whole_selector(arg: AstNode | None) -> bool:
        """True when *arg* is the literal `0` (Excel whole row/column selector)."""
        return isinstance(arg, NumberNode) and arg.value == 0

    def _index_range_result_is_scalar(self, range_node: RangeNode, node: FunctionCallNode) -> bool:
        """True when INDEX(range, ...) resolves to a single cell (not a row/column slice)."""
        _, r1, c1, r2, c2 = self._range_coords(range_node.start, range_node.end)
        nrows = r2 - r1 + 1
        ncols = c2 - c1 + 1

        row_arg = node.args[1] if len(node.args) > 1 else None
        col_arg = node.args[2] if len(node.args) > 2 else None
        row_omitted = row_arg is None or isinstance(row_arg, EmptyArgNode)
        col_omitted = col_arg is None or isinstance(col_arg, EmptyArgNode)

        if row_omitted and col_omitted:
            return nrows == 1 and ncols == 1
        if row_omitted:
            if self._index_arg_is_whole_selector(col_arg):
                return nrows == 1 and ncols == 1
            return nrows == 1
        if col_omitted:
            if self._index_arg_is_whole_selector(row_arg):
                return nrows == 1 and ncols == 1
            return ncols == 1
        if self._index_arg_is_whole_selector(row_arg) or self._index_arg_is_whole_selector(col_arg):
            return nrows == 1 and ncols == 1
        return True

    def _emit_range_ref_tuple(self, range_node: RangeNode) -> str:
        """Emit a range as an (sheet, r1, c1, r2, c2) tuple for xl_index_ref / xl_offset."""
        base_sheet, r1, c1, r2, c2 = self._range_coords(range_node.start, range_node.end)
        self._offset_runtime_sheets.add(base_sheet)
        return f"({repr(base_sheet)}, {r1}, {c1}, {r2}, {c2})"

    def _emit_index_scalar_range(self, node: FunctionCallNode) -> str:
        """Emit INDEX over a literal range when the result is a single cell."""
        base = node.args[0]
        assert isinstance(base, RangeNode)
        base_ref_info = self._emit_range_ref_tuple(base)
        row_expr = (
            "None"
            if len(node.args) < 2 or isinstance(node.args[1], EmptyArgNode)
            else self._emit_ast_child(node.args[1])
        )
        col_expr = (
            "None"
            if len(node.args) < 3 or isinstance(node.args[2], EmptyArgNode)
            else self._emit_ast_child(node.args[2])
        )
        self._needs_offset_runtime = True
        self._needs_index_ref_runtime = True
        expr = f"xl_offset(ctx, xl_index_ref({base_ref_info}, {row_expr}, {col_expr}), 0.0, 0.0)"
        return self._hoist_return_expr(expr)

    def _range_coords(self, start: str, end: str) -> tuple[str, int, int, int, int]:
        """Parse a range into (sheet, start_row, start_col, end_row, end_col).

        Uses Excel semantics: start sheet applies to the whole range.
        """
        start_sheet, start_cell = parse_address(start)
        _, end_cell = parse_address(end)

        start_col_str, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col_str, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell)

        start_col = fastpyxl.utils.cell.column_index_from_string(start_col_str)
        end_col = fastpyxl.utils.cell.column_index_from_string(end_col_str)

        r1, r2 = (start_row, end_row) if start_row <= end_row else (end_row, start_row)
        c1, c2 = (start_col, end_col) if start_col <= end_col else (end_col, start_col)
        return (start_sheet, r1, c1, r2, c2)

    def _range_node_is_single_cell(self, node: RangeNode, *, anchor: str | None = None) -> bool:
        """True when `node` spans exactly one cell (e.g. `A1:A1`)."""
        start = resolve_cell_ref(node.start_ref, anchor)
        end = resolve_cell_ref(node.end_ref, anchor)
        _, r1, c1, r2, c2 = self._range_coords(start, end)
        return r1 == r2 and c1 == c2

    def _range_cell_count(self, start: str, end: str) -> int:
        _, r1, c1, r2, c2 = self._range_coords(start, end)
        return (r2 - r1 + 1) * (c2 - c1 + 1)

    def _max_array_extent_in_ast(self, node: AstNode) -> int:
        """Return the largest array operand size referenced in *node*."""
        if isinstance(node, RangeNode):
            return self._range_cell_count(node.start, node.end)
        if isinstance(node, BinaryOpNode):
            return max(
                self._max_array_extent_in_ast(node.left),
                self._max_array_extent_in_ast(node.right),
            )
        if isinstance(node, UnaryOpNode):
            return self._max_array_extent_in_ast(node.operand)
        if isinstance(node, FunctionCallNode):
            extent = 1
            for arg in node.args:
                extent = max(extent, self._max_array_extent_in_ast(arg))
            return extent
        return 1

    def _ast_needs_array_operator_branch(self, node: SkeletonNode) -> bool:
        """Return whether a subtree can *evaluate to* a range/array at runtime.

        Only array-producing nodes require the operator broadcast branch: ranges,
        multi-cell `OFFSET`, non-scalar `INDEX` slices, and the pass-through
        functions (`IF`/`IFERROR`/`IFNA`/`CHOOSE`) when a returned branch is
        itself an array. Scalar-returning functions (e.g. `SUM`, `MATCH`,
        `VLOOKUP`) never yield arrays even when their arguments contain ranges,
        so they take the inlined scalar path without a guard. A 1x1 range is a
        scalar cell read, not an array producer.
        """
        if isinstance(node, RangeNode):
            return not self._range_node_is_single_cell(node)
        if isinstance(node, AddressHoleNode):
            return node.kind in {"RANGE", "WHOLE_COL", "WHOLE_ROW"}
        if isinstance(node, BinaryOpNode):
            return self._ast_needs_array_operator_branch(
                node.left
            ) or self._ast_needs_array_operator_branch(node.right)
        if isinstance(node, UnaryOpNode):
            return self._ast_needs_array_operator_branch(node.operand)
        if isinstance(node, FunctionCallNode):
            upper = normalize_excel_function_name(node.name)
            if upper == "OFFSET":
                return not self._can_offset_be_static(node) or self._static_offset_is_multicell(
                    node
                )
            if upper == "INDEX" and node.args and isinstance(node.args[0], RangeNode):
                return not self._index_range_result_is_scalar(node.args[0], node)
            if upper in {"IFERROR", "IFNA"}:
                return any(self._ast_needs_array_operator_branch(arg) for arg in node.args)
            if upper in {"IF", "CHOOSE"}:
                # The result is one of the value branches (args[1:]); the
                # condition/index (args[0]) does not affect array-ness.
                return any(self._ast_needs_array_operator_branch(arg) for arg in node.args[1:])
            return False
        return False

    def _note_operators_fastpath_from_ast(self, node: AstNode) -> None:
        if self._max_array_extent_in_ast(node) >= MIN_OPERATOR_FASTPATH_CELLS:
            self._needs_operators_fastpath = True

    def _emit_offset_static(
        self, base_address: str, rows: int, cols: int, height: int, width: int
    ) -> str:
        """Emit statically resolved OFFSET as direct cell/range reference."""
        base_sheet, base_cell = parse_address(base_address)
        base_col_str, base_row = fastpyxl.utils.cell.coordinate_from_string(base_cell)
        base_col = fastpyxl.utils.cell.column_index_from_string(base_col_str)

        # Compute target position
        target_row = base_row + rows
        target_col = base_col + cols

        if target_row < 1 or target_col < 1:
            # Invalid reference
            return "xl_raise(XlError.REF)"

        target_col_str = fastpyxl.utils.cell.get_column_letter(target_col)

        if height == 1 and width == 1:
            # Single cell reference
            target_addr = f"{quote_sheet_if_needed(base_sheet)}!{target_col_str}{target_row}"
            return self._emit_cell_eval(target_addr)
        else:
            # Range reference - emit as a lazy range like _emit_range does
            end_row = target_row + height - 1
            end_col = target_col + width - 1
            end_col_str = fastpyxl.utils.cell.get_column_letter(end_col)

            start_addr = f"{quote_sheet_if_needed(base_sheet)}!{target_col_str}{target_row}"
            end_addr = f"{quote_sheet_if_needed(base_sheet)}!{end_col_str}{end_row}"
            return self._emit_range_address(start_addr, end_addr)

    def _emit_offset_dynamic(
        self,
        ref_node: AstNode,
        rows_node: AstNode,
        cols_node: AstNode,
        height_node: AstNode | None,
        width_node: AstNode | None,
    ) -> str:
        """Emit dynamic OFFSET that resolves at runtime."""
        # For dynamic OFFSET, we need to pass the base reference info
        if isinstance(ref_node, CellRefNode):
            base_sheet, base_cell = parse_address(ref_node.address)
            self._offset_runtime_sheets.add(base_sheet)
            base_col_str, base_row = fastpyxl.utils.cell.coordinate_from_string(base_cell)
            base_col = fastpyxl.utils.cell.column_index_from_string(base_col_str)
            ref_info = f"({repr(base_sheet)}, {base_row}, {base_col})"
        elif isinstance(ref_node, RangeNode):
            base_sheet, r1, c1, r2, c2 = self._range_coords(ref_node.start, ref_node.end)
            self._offset_runtime_sheets.add(base_sheet)
            ref_info = f"({repr(base_sheet)}, {r1}, {c1}, {r2}, {c2})"
        elif isinstance(ref_node, FunctionCallNode) and ref_node.name.upper() == "INDEX":
            if len(ref_node.args) < 1:
                return "xl_raise(XlError.VALUE)"
            base = ref_node.args[0]
            if isinstance(base, CellRefNode):
                base_sheet, base_cell = parse_address(base.address)
                self._offset_runtime_sheets.add(base_sheet)
                base_col_str, base_row = fastpyxl.utils.cell.coordinate_from_string(base_cell)
                base_col = fastpyxl.utils.cell.column_index_from_string(base_col_str)
                base_ref_info = f"({repr(base_sheet)}, {base_row}, {base_col})"
            elif isinstance(base, RangeNode):
                base_sheet, r1, c1, r2, c2 = self._range_coords(base.start, base.end)
                self._offset_runtime_sheets.add(base_sheet)
                base_ref_info = f"({repr(base_sheet)}, {r1}, {c1}, {r2}, {c2})"
            else:
                return "xl_raise(XlError.REF)"

            row_expr = (
                "None"
                if len(ref_node.args) < 2 or isinstance(ref_node.args[1], EmptyArgNode)
                else self._emit_ast_child(ref_node.args[1])
            )
            col_expr = (
                "None"
                if len(ref_node.args) < 3 or isinstance(ref_node.args[2], EmptyArgNode)
                else self._emit_ast_child(ref_node.args[2])
            )
            self._needs_index_ref_runtime = True
            ref_info = f"xl_index_ref({base_ref_info}, {row_expr}, {col_expr})"
        else:
            # If reference is not a simple cell, we can't handle it
            return "xl_raise(XlError.REF)"

        rows_expr = self._emit_ast_child(rows_node)
        cols_expr = self._emit_ast_child(cols_node)
        height_expr = "None" if height_node is None else self._emit_ast_child(height_node)
        width_expr = "None" if width_node is None else self._emit_ast_child(width_node)

        expr = f"xl_offset(ctx, {ref_info}, {rows_expr}, {cols_expr}, {height_expr}, {width_expr})"
        return self._hoist_return_expr(expr)

    def _emit_offset_ref(self, node: FunctionCallNode) -> str:
        if len(node.args) < 3:
            return "xl_raise(XlError.VALUE)"

        ref_node = node.args[0]
        rows_node = node.args[1]
        cols_node = node.args[2]
        height_node = node.args[3] if len(node.args) > 3 else None
        width_node = node.args[4] if len(node.args) > 4 else None

        if isinstance(ref_node, CellRefNode):
            base_sheet, base_cell = parse_address(ref_node.address)
            base_col_str, base_row = fastpyxl.utils.cell.coordinate_from_string(base_cell)
            base_col = fastpyxl.utils.cell.column_index_from_string(base_col_str)
            ref_info = f"({repr(base_sheet)}, {base_row}, {base_col})"
        elif isinstance(ref_node, RangeNode):
            base_sheet, r1, c1, r2, c2 = self._range_coords(ref_node.start, ref_node.end)
            ref_info = f"({repr(base_sheet)}, {r1}, {c1}, {r2}, {c2})"
        elif isinstance(ref_node, FunctionCallNode) and ref_node.name.upper() == "INDEX":
            if len(ref_node.args) < 1:
                return "xl_raise(XlError.VALUE)"
            base = ref_node.args[0]
            if isinstance(base, CellRefNode):
                base_sheet, base_cell = parse_address(base.address)
                base_col_str, base_row = fastpyxl.utils.cell.coordinate_from_string(base_cell)
                base_col = fastpyxl.utils.cell.column_index_from_string(base_col_str)
                base_ref_info = f"({repr(base_sheet)}, {base_row}, {base_col})"
            elif isinstance(base, RangeNode):
                base_sheet, r1, c1, r2, c2 = self._range_coords(base.start, base.end)
                base_ref_info = f"({repr(base_sheet)}, {r1}, {c1}, {r2}, {c2})"
            else:
                return "xl_raise(XlError.REF)"

            row_expr = (
                "None"
                if len(ref_node.args) < 2 or isinstance(ref_node.args[1], EmptyArgNode)
                else self._emit_ast_child(ref_node.args[1])
            )
            col_expr = (
                "None"
                if len(ref_node.args) < 3 or isinstance(ref_node.args[2], EmptyArgNode)
                else self._emit_ast_child(ref_node.args[2])
            )
            self._needs_index_ref_runtime = True
            ref_info = f"xl_index_ref({base_ref_info}, {row_expr}, {col_expr})"
        else:
            return "xl_raise(XlError.REF)"

        rows_expr = self._emit_ast_child(rows_node)
        cols_expr = self._emit_ast_child(cols_node)
        height_expr = "None" if height_node is None else self._emit_ast_child(height_node)
        width_expr = "None" if width_node is None else self._emit_ast_child(width_node)

        expr = f"xl_offset_ref({ref_info}, {rows_expr}, {cols_expr}, {height_expr}, {width_expr})"
        return self._hoist_return_expr(expr)

    def _start_return_unpack(self) -> list[str]:
        if not self._unpack_return:
            return []
        self._return_unpack_state = _ReturnUnpackState(statements=[])
        self._return_unpack_stack = [_ReturnUnpackFrame(lazy=False, nested=False)]
        return self._return_unpack_state.statements

    def _stop_return_unpack(self) -> None:
        self._return_unpack_state = None
        self._return_unpack_stack = []

    @staticmethod
    def _format_return_lines(unpack_stmts: list[str], expr: str) -> list[str]:
        lines = [f"    {stmt}" for stmt in unpack_stmts]
        lines.append(f"    return {expr}")
        return lines

    def _emit_formula_body_lines(self, ast: SkeletonNode) -> list[str]:
        """Emit statement/return lines for a formula cell or alias body."""
        unpack_stmts = self._start_return_unpack()
        try:
            expr = self._emit_ast(ast)
        finally:
            self._stop_return_unpack()
        return self._format_return_lines(unpack_stmts, expr)

    @staticmethod
    def _shape_helper_eligible(skeleton: SkeletonNode) -> bool:
        """Return whether `skeleton` can be emitted as a shared param helper.

        Helpers bind addresses at call time, so emit must not depend on a
        concrete `CellRefNode` / `RangeNode`. OFFSET, INDEX, and ROW/COLUMN/
        COLUMNS inspect those nodes (or the formula cell address), so they
        stay on the specialized filled-AST body.
        """
        for hole in iter_address_holes(skeleton):
            if hole.kind not in {"CELL", "RANGE"}:
                return False

        def walk(node: object) -> bool:
            if isinstance(node, FunctionCallNode):
                if normalize_excel_function_name(node.name) in _SHAPE_HELPER_REF_FUNCS:
                    return False
                return all(walk(arg) for arg in node.args)
            if isinstance(node, BinaryOpNode):
                return walk(node.left) and walk(node.right)
            if isinstance(node, UnaryOpNode):
                return walk(node.operand)
            return True

        return walk(skeleton)

    def _shape_params_include_scalar_range(
        self, params: tuple[AddressLeaf, ...], *, host: str
    ) -> bool:
        """Return whether any RANGE param is a 1x1 range (inlined as a scalar)."""
        return any(
            isinstance(leaf, RangeNode) and self._range_node_is_single_cell(leaf, anchor=host)
            for leaf in params
        )

    def _plan_shape_helpers(self, formula_addresses: Sequence[str]) -> None:
        """Record profitable shape helper names for this generate() pass."""
        self._shape_helper_names.clear()
        if self._reader_index is not None:
            return
        table = getattr(self.graph, "formula_shapes", None)
        if table is None:
            return
        counts: dict[str, int] = {}
        scalar_range_shapes: set[str] = set()
        for address in formula_addresses:
            node = self.graph.get_node(address)
            if node is None or not _node_has_formula(node):
                continue
            found = table.lookup(address)
            if found is None:
                continue
            shape_key, skeleton, params = found
            if not self._shape_helper_eligible(skeleton):
                continue
            if self._shape_params_include_scalar_range(params, host=address):
                scalar_range_shapes.add(shape_key)
            counts[shape_key] = counts.get(shape_key, 0) + 1
        profitable = sorted(
            key for key, n in counts.items() if n >= 2 and key not in scalar_range_shapes
        )
        self._shape_helper_names = {
            shape_key: f"_shape_{index}" for index, shape_key in enumerate(profitable)
        }

    def _emit_shape_helpers(self) -> list[str]:
        """Emit shared per-shape helpers referenced by formula cell wrappers."""
        table = getattr(self.graph, "formula_shapes", None)
        if table is None or not self._shape_helper_names:
            return []
        lines: list[str] = ["# --- Shared formula-shape helpers ---", ""]
        for shape_key, helper_name in self._shape_helper_names.items():
            skeleton = table.shapes[shape_key]
            n_holes = len(list(iter_address_holes(skeleton)))
            params = ", ".join(["ctx", *[f"p{i}" for i in range(n_holes)]])
            lines.append(f"def {helper_name}({params}):")
            self._temp_var_counter = 0
            prev_cell = self._formula_cell_address
            self._formula_cell_address = None
            try:
                body = self._emit_formula_body_lines(skeleton)
            finally:
                self._formula_cell_address = prev_cell
            lines.extend(body)
            lines.append("")
            lines.append("")
        return lines

    def _emit_shape_helper_call(
        self, helper_name: str, params: tuple[AddressLeaf, ...], host_address: str
    ) -> list[str]:
        encoded = ", ".join(repr(resolve_address_leaf(leaf, host_address)) for leaf in params)
        if encoded:
            return [f"    return {helper_name}(ctx, {encoded})"]
        return [f"    return {helper_name}(ctx)"]

    def _emit_cell(self, address: str) -> str:
        """Emit a Python function for a single formula cell.

        Args:
            address: Sheet-qualified cell address (e.g., 'Sheet1!A1')

        Returns:
            Python function definition as a string.
        """
        normalized = normalize_address(address)
        func_name = address_to_python_name(normalized)
        node = self.graph.get_node(normalized)

        if node is None or not _node_has_formula(node):
            raise ValueError(f"Not a formula cell: {normalized}")

        lines: list[str] = []
        lines.append(f"def {func_name}(ctx):")
        # Prefer the raw workbook text when it was stored; otherwise document the
        # AST-derived (or stored) normalized formula.
        shown_formula = display_formula(node) or "="
        doc = f"Formula: {shown_formula}".replace("'''", "\\'''")
        if doc[-1] not in ".?!":
            doc = f"{doc}."
        lines.append(f"    '''{doc}'''")
        table = getattr(self.graph, "formula_shapes", None)
        if table is not None and self._shape_helper_names:
            found = table.lookup(normalized)
            if found is not None:
                shape_key, _skeleton, params = found
                helper_name = self._shape_helper_names.get(shape_key)
                if helper_name is not None:
                    lines.extend(self._emit_shape_helper_call(helper_name, params, normalized))
                    return "\n".join(lines)
        # Reset temp var counter for each cell to keep variable names short
        self._temp_var_counter = 0
        ast = self._get_or_parse_ast(normalized)
        assert ast is not None
        prev_cell = self._formula_cell_address
        self._formula_cell_address = normalized
        try:
            lines.extend(self._emit_formula_body_lines(ast))
        finally:
            self._formula_cell_address = prev_cell

        return "\n".join(lines)

    @staticmethod
    def _is_xlerror_type_expr(node: py_ast.AST) -> bool:
        if isinstance(node, py_ast.Name):
            return node.id == "XlError"
        if isinstance(node, py_ast.Tuple):
            return any(CodeGenerator._is_xlerror_type_expr(elt) for elt in node.elts)
        return False

    @classmethod
    def _cell_function_has_xlerror_isinstance(cls, node: py_ast.FunctionDef) -> bool:
        for child in py_ast.walk(node):
            if not isinstance(child, py_ast.Call):
                continue
            if not isinstance(child.func, py_ast.Name) or child.func.id != "isinstance":
                continue
            if len(child.args) < 2:
                continue
            if cls._is_xlerror_type_expr(child.args[1]):
                return True
        return False

    @classmethod
    def _assert_raise_only_cell_boundary(cls, cell_code_lines: list[str]) -> None:
        """Reject generated formula cells that inspect `XlError` sentinels."""
        module = py_ast.parse("\n".join(cell_code_lines))
        offenders = [
            node.name
            for node in module.body
            if isinstance(node, py_ast.FunctionDef)
            and node.name.startswith("cell_")
            and cls._cell_function_has_xlerror_isinstance(node)
        ]
        if offenders:
            joined = ", ".join(offenders)
            raise ValueError(
                "Generated formula cell functions must not inspect XlError sentinels "
                f"with isinstance(): {joined}"
            )

    def _collect_dependencies(self, address: str) -> list[str]:
        """Collect all cell addresses that a cell depends on (recursively).

        Cycles are allowed in the dependency graph. Excel permits circular
        references when broken by conditional evaluation (IF, IFERROR, etc.).
        The generated code handles cycles at runtime via EvalContext tracking
        and lazy evaluation of function arguments.

        Missing cells (referenced but not in graph) are included in the output
        so stub functions can be generated for them.

        Args:
            address: Starting cell address.

        Returns:
            List of all dependent cell addresses in dependency order.
            Includes missing cells (not in graph) that are referenced by formulas.
        """
        visited: set[str] = set()
        in_progress: set[str] = set()  # Currently being visited (for cycle detection)
        order: list[str] = []

        def visit(addr: str) -> None:
            # Normalize address to match Node.key format
            addr = normalize_address(addr)

            # Skip if already fully visited or currently being visited (cycle)
            if addr in visited or addr in in_progress:
                return

            in_progress.add(addr)

            node = self.graph.get_node(addr)
            if node is None:
                # Cell not in graph - still add to order so we generate a stub
                order.append(addr)
                in_progress.discard(addr)
                visited.add(addr)
                return

            # If it's a formula, parse and find cell references
            if _node_has_formula(node):
                ast = self._get_or_parse_ast(addr)
                assert ast is not None
                deps = self._extract_cell_refs(ast)
                for dep in deps:
                    visit(dep)

            order.append(addr)
            in_progress.discard(addr)
            visited.add(addr)

        visit(address)
        return order

    def _extract_cell_refs(self, node: AstNode) -> list[str]:
        """Extract all cell references from an AST node."""
        refs: list[str] = []

        if isinstance(node, CellRefNode):
            refs.append(node.address)
        elif isinstance(node, RangeNode):
            refs.extend(self._range_addresses(node.start, node.end))
        elif isinstance(node, BinaryOpNode):
            refs.extend(self._extract_cell_refs(node.left))
            refs.extend(self._extract_cell_refs(node.right))
        elif isinstance(node, UnaryOpNode):
            refs.extend(self._extract_cell_refs(node.operand))
        elif isinstance(node, FunctionCallNode):
            # For OFFSET that can be statically resolved, extract target cells
            if node.name.upper() == "OFFSET" and self._can_offset_be_static(node):
                refs.extend(self._extract_offset_target_refs(node))
            for arg in node.args:
                refs.extend(self._extract_cell_refs(arg))

        return refs

    def _extract_offset_target_refs(self, node: FunctionCallNode) -> list[str]:
        """Extract target cell references from a statically resolvable OFFSET."""
        ref_node = node.args[0]
        rows_node = node.args[1]
        cols_node = node.args[2]
        height_node = node.args[3] if len(node.args) > 3 else None
        width_node = node.args[4] if len(node.args) > 4 else None

        if isinstance(ref_node, CellRefNode):
            base_address = ref_node.address
            base_h, base_w = (1, 1)
        elif isinstance(ref_node, RangeNode):
            base_address = ref_node.start
            base_h, base_w = self._offset_base_shape(ref_node)
        else:
            return []

        base_sheet, base_cell = parse_address(base_address)
        base_col_str, base_row = fastpyxl.utils.cell.coordinate_from_string(base_cell)
        base_col = fastpyxl.utils.cell.column_index_from_string(base_col_str)

        rows = int(self._get_constant_number(rows_node))
        cols = int(self._get_constant_number(cols_node))
        height = int(self._get_constant_number(height_node)) if height_node is not None else base_h
        width = int(self._get_constant_number(width_node)) if width_node is not None else base_w

        target_row = base_row + rows
        target_col = base_col + cols

        if target_row < 1 or target_col < 1:
            return []

        refs = []
        for r in range(target_row, target_row + height):
            for c in range(target_col, target_col + width):
                col_str = fastpyxl.utils.cell.get_column_letter(c)
                refs.append(f"{quote_sheet_if_needed(base_sheet)}!{col_str}{r}")

        return refs

    def _extract_xl_functions(self, node: AstNode) -> set[str]:
        """Extract all xl_* function names and markers used in an AST node.

        Special markers:
        - "XlError": XlError enum is needed (e.g., error literals like #N/A)
        """
        funcs: set[str] = set()

        if isinstance(node, ErrorNode):
            # Error literals raise via xl_raise and reference the XlError enum
            funcs.add("XlError")
            funcs.add("xl_raise")
        elif isinstance(node, RangeNode):
            # Multi-cell ranges emit lazy xl_range(ctx, ...); 1x1 collapses to
            # a cell read (xl_cell / xl_eval), discovered from emitted bodies.
            if not self._range_node_is_single_cell(node):
                funcs.add("xl_range")
        elif isinstance(node, FunctionCallNode):
            upper_name = normalize_excel_function_name(node.name)

            # NA() emits a raising error literal (xl_raise(XlError.NA)).
            if upper_name == "NA":
                funcs.add("XlError")
                funcs.add("xl_raise")
            # IF, IFERROR, CHOOSE are special - emitted as native Python conditionals
            elif upper_name == "IF":
                funcs.add("xl_bool")
            elif upper_name == "IFERROR":
                funcs.add("XlError")
                funcs.add("xl_iferror")
            elif upper_name == "IFNA":
                funcs.add("XlError")
                funcs.add("xl_ifna")
            elif upper_name == "CHOOSE":
                funcs.add("XlError")
                funcs.add("xl_int")
                funcs.add("xl_raise")
            elif upper_name == "ROW":
                funcs.add("xl_row")
                if node.args:
                    ref = node.args[0]
                    if isinstance(ref, FunctionCallNode) and ref.name.upper() == "OFFSET":
                        funcs.add("xl_offset_ref")
                        for off_arg in ref.args:
                            funcs.update(self._extract_xl_functions(off_arg))
                    else:
                        funcs.update(self._extract_xl_functions(ref))
            elif upper_name in {"COLUMN", "COLUMNS"}:
                funcs.add("xl_column" if upper_name == "COLUMN" else "xl_columns")
                if node.args:
                    ref = node.args[0]
                    if isinstance(ref, FunctionCallNode) and ref.name.upper() == "OFFSET":
                        funcs.add("xl_offset_ref")
                        for off_arg in ref.args:
                            funcs.update(self._extract_xl_functions(off_arg))
                    else:
                        funcs.update(self._extract_xl_functions(ref))
            # OFFSET is special - only add xl_offset if it can't be statically resolved
            elif upper_name == "OFFSET":
                if not self._can_offset_be_static(node):
                    funcs.add("xl_offset")
                elif self._static_offset_is_multicell(node):
                    funcs.add("xl_range")
            elif (
                upper_name == "INDEX"
                and node.args
                and isinstance(node.args[0], RangeNode)
                and self._index_range_result_is_scalar(node.args[0], node)
            ):
                funcs.add("xl_offset")
                funcs.add("xl_index_ref")
            else:
                funcs.add(excel_func_to_python(node.name))

            skip_index_range_arg = (
                upper_name == "INDEX"
                and node.args
                and isinstance(node.args[0], RangeNode)
                and self._index_range_result_is_scalar(node.args[0], node)
            )
            for i, arg in enumerate(node.args):
                if skip_index_range_arg and i == 0:
                    continue
                funcs.update(self._extract_xl_functions(arg))
        elif isinstance(node, BinaryOpNode):
            needs_array = self._ast_needs_array_operator_branch(node)
            if needs_array:
                funcs.add("xl_is_array")
            if node.op == "&":
                if needs_array:
                    funcs.add("xl_map_concat")
                funcs.add("to_string")
            elif node.op in _ARITHMETIC_OPS:
                funcs.add("xl_number")
                if needs_array:
                    funcs.add("xl_map_arithmetic")
                if node.op == "^":
                    funcs.add("xl_pow_numbers")
                if node.op == "/":
                    funcs.add("xl_raise")
                    funcs.add("XlError")
            elif node.op in _COMPARE_OPS:
                funcs.add("xl_compare")
                if needs_array:
                    funcs.add("xl_map_compare")
            funcs.update(self._extract_xl_functions(node.left))
            funcs.update(self._extract_xl_functions(node.right))
        elif isinstance(node, UnaryOpNode):
            funcs.add("xl_number")
            if self._ast_needs_array_operator_branch(node):
                funcs.add("xl_is_array")
                funcs.add("xl_map_unary")
            funcs.update(self._extract_xl_functions(node.operand))

        return funcs

    def generate(
        self,
        targets: Sequence[str] | None = None,
        *,
        constant_types: set[str] | None = None,
        constant_ranges: Sequence[str] | None = None,
        constant_blanks: bool = False,
        input_ranges: Sequence[str] | None = None,
        blank_ranges: Sequence[str] | None = None,
        series_bindings: WorkbookSeriesBindings | None = None,
        bindings_workbook: Path | str | None = None,
        series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
        docstring_renderer: SeriesDocstringRendererSpec = "google",
        include_compute_all: bool | None = None,
    ) -> str:
        """Generate standalone Python code for target cells.

        When `graph.formula_shapes` is set, this pass may emit shared
        per-shape helpers. The overlay is read at generate time (not an init
        snapshot). Missing shapes fall back to per-node AST.

        Args:
            targets: List of target cell addresses to compute.
            constant_types: Cell value kinds emitted as typed constants in generated code.
            constant_ranges: Sheet-qualified ranges whose cells are emitted as constants.
            constant_blanks: When True, blank cells in constant ranges become `None`.
            input_ranges: Sheet-qualified ranges whose leaf cells are treated as inputs.
                When a cell would otherwise be a constant, `input_ranges` take precedence.
            blank_ranges: Sheet-qualified rectangles whose cells are omitted from the graph
                but resolve as empty (`None`) at runtime; must match builder/evaluator.
            series_bindings: Optional workbook binding manifest; when set with
                `bindings_workbook`, emits `set_*` functions that accept Records.
            bindings_workbook: Path to the `.xlsx` file used to resolve binds.
            series_docstring_callback: Optional registered callback name for structured
                docstrings on generated `set_*` and series output `compute_*` functions.
            docstring_renderer: Built-in renderer name or custom renderer object/callable
                for structured series-binding docstrings (`plain`, `rst`,
                `google`, `numpy`).
            include_compute_all: Control emission of public `compute_all`. `True` always
                emits it, `False` never emits it, and `None` (default) omits it when
                every export target is covered by an output series binding.

        Returns:
            Standalone Python source code as a string.
        """
        normalized_targets = self._resolve_targets(targets)

        parts = self._generate_parts(
            normalized_targets,
            dependency_targets=normalized_targets,
            constant_types=constant_types,
            constant_ranges=constant_ranges,
            constant_blanks=constant_blanks,
            input_ranges=input_ranges,
            blank_ranges=blank_ranges,
            series_bindings=series_bindings,
            bindings_workbook=bindings_workbook,
        )
        runtime_code = parts["runtime_code"]
        cell_code_lines = parts["cell_code_lines"]
        _formula_cells = parts["formula_cells"]
        _all_cells = parts["all_cells"]
        normalized_targets = parts["targets"]
        series_public_addresses = self._series_binding_public_addresses(
            series_bindings,
            bindings_workbook,
        )
        public_addresses = frozenset(normalized_targets) | series_public_addresses

        # Combine: runtime + inputs + formulas + entry point
        if parts["has_constants"]:
            runtime_code = _ensure_mapping_proxy_import(runtime_code)
        lines: list[str] = [runtime_code, ""]

        lines.extend(parts["inputs_block_lines"])
        if parts["has_constants"]:
            lines.append("")
            if parts["constants_block_lines"]:
                lines.extend(parts["constants_block_lines"])
            else:
                lines.append("CONSTANTS = MappingProxyType({})")
        lines.append("")
        lines.append("")

        # Export formula cell implementations and a resolver.
        lines.append("# --- Formula cell functions ---\n")
        lines.extend(cell_code_lines)
        alias_lines = self._emit_projection_alias_lines(_all_cells, public_addresses)
        lines.extend(alias_lines)
        lines.extend(
            self._emit_resolver_lines(parts["blank_rects"] if parts["blank_rects"] else None)
        )

        # Generate entry point helpers
        lines.append("def make_context(inputs: dict[str, object] | None = None) -> EvalContext:")
        lines.append('    """Create an EvalContext with merged inputs."""')
        if parts["has_constants"]:
            lines.append("    merged = prepare_context_inputs(DEFAULT_INPUTS, CONSTANTS, inputs)")
        else:
            lines.append("    merged = prepare_context_inputs(DEFAULT_INPUTS, None, inputs)")
        lines.append(
            "    return EvalContext("
            "inputs=merged, resolver=_resolve_formula, "
            f"iterative_enabled={bool(self._iterate_enabled)}, "
            f"iterate_count={int(self._iterate_count)}, "
            f"iterate_delta={float(self._iterate_delta)!r})"
        )
        lines.append("")
        lines.append("")
        series_setter_names: list[str] = []
        series_reader_names: list[str] = []
        series_compute_names: list[str] = []
        reader_leaves: dict[str, dict[str, object]] | None = None
        reader_ranges: dict[str, dict[str, object]] | None = None
        if series_bindings is not None:
            if bindings_workbook is None:
                raise ValueError("bindings_workbook is required when series_bindings is set")
            (
                series_setter_names,
                series_reader_names,
                series_compute_names,
            ) = self._series_binding_public_names(series_bindings)
            lines.extend(
                self._emit_series_binding_setters(
                    series_bindings,
                    bindings_workbook,
                    export_addresses=_all_cells,
                    public_addresses=series_public_addresses,
                    series_docstring_callback=series_docstring_callback,
                    docstring_renderer=docstring_renderer,
                )
            )
            lines.append("")
            if self._reader_index is not None:
                from excel_grapher.series_bindings.reader_index import (
                    reader_index_as_discovery_dicts,
                )

                reader_leaves, reader_ranges = reader_index_as_discovery_dicts(self._reader_index)
            else:
                reader_leaves, reader_ranges = self._series_binding_reader_discovery(
                    cast("DependencyGraph", self._public_graph()),
                    series_bindings,
                    workbook=bindings_workbook,
                    export_addresses=_all_cells,
                )
        lines.extend(
            self._emit_series_binding_discovery_lines(
                series_setter_names,
                series_compute_names,
                self._series_binding_groups_manifest(series_bindings),
                reader_names=series_reader_names,
                reader_leaves=reader_leaves,
                reader_ranges=reader_ranges,
            )
        )
        emit_compute_all = self._should_emit_compute_all(
            normalized_targets,
            series_bindings=series_bindings,
            bindings_workbook=bindings_workbook,
            export_addresses=_all_cells,
            include_compute_all=include_compute_all,
        )
        if emit_compute_all:
            lines.append("")
            lines.append("")
            lines.extend(self._emit_compute_all_block(normalized_targets))

        return "\n".join(lines)

    def generate_modules(
        self,
        targets: Sequence[str] | None = None,
        *,
        constant_types: set[str] | None = None,
        constant_ranges: Sequence[str] | None = None,
        constant_blanks: bool = False,
        input_ranges: Sequence[str] | None = None,
        blank_ranges: Sequence[str] | None = None,
        series_bindings: WorkbookSeriesBindings | None = None,
        bindings_workbook: Path | str | None = None,
        series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
        docstring_renderer: SeriesDocstringRendererSpec = "google",
        address_helpers: Mapping[str, OutputHelperSpec] | None = None,
        include_compute_all: bool | None = None,
    ) -> dict[str, str]:
        """Generate a multi-module Python package for target cells.

        Returns a mapping of module filenames to file contents. Callers choose how to
        lay out files on disk (e.g. a package directory).

        The generated package has five flat files:
        - __init__.py: exports public API names and DEFAULT_INPUTS
        - api.py: make_context, optional compute_all, and series-binding set_* / read_* / compute_*
        - data.py: DEFAULT_INPUTS and CONSTANTS
        - runtime.py: embedded Excel runtime (emit_runtime)
        - internals.py: formula cell functions + resolver dispatch

        When series bindings declare input series, the package also includes:
        - `_readers.py`: leaf maps and `read_*` duals (imported by `api` and `internals`)
        - `_api_helpers.py`: coercion helpers for setters

        When series bindings declare output series, the package also includes:
        - `_output_leaves.py`: `_OUTPUT_LEAVES_*` tables imported by `api`

        Public `read_*` helpers are defined in `_readers.py` and re-exported from
        `api.py` (via `__all__`) so `from <pkg>.api import read_foo` works.

        `compute_all` is omitted by default when every export target is covered by an
        output series binding. Pass `include_compute_all=True` to keep it, or
        `include_compute_all=False` to omit it unconditionally.
        """
        normalized_targets = self._resolve_targets(targets)

        parts = self._generate_parts(
            normalized_targets,
            dependency_targets=normalized_targets,
            constant_types=constant_types,
            constant_ranges=constant_ranges,
            constant_blanks=constant_blanks,
            input_ranges=input_ranges,
            blank_ranges=blank_ranges,
            series_bindings=series_bindings,
            bindings_workbook=bindings_workbook,
        )
        runtime_code = parts["runtime_code"]
        cell_code_lines = parts["cell_code_lines"]
        _formula_cells = parts["formula_cells"]
        _all_cells = parts["all_cells"]
        normalized_targets = parts["targets"]
        series_public_addresses = self._series_binding_public_addresses(
            series_bindings,
            bindings_workbook,
        )
        public_addresses = frozenset(normalized_targets) | series_public_addresses

        data_lines_out: list[str] = [
            "from __future__ import annotations",
            "",
            _MAPPING_PROXY_IMPORT,
            "",
            "# --- Default inputs (leaf cells) ---",
            *parts["inputs_block_lines"][1:],  # drop the comment already included above
            "",
        ]
        if parts["constants_block_lines"]:
            data_lines_out.extend(parts["constants_block_lines"])
        else:
            data_lines_out.append("CONSTANTS = MappingProxyType({})")
        data_lines_out.append("")
        data_py = "\n".join(data_lines_out).rstrip() + "\n"

        runtime_py = runtime_code.rstrip() + "\n"

        alias_lines = self._emit_projection_alias_lines(_all_cells, public_addresses)
        internals_import_names = self._internals_runtime_import_names(
            parts["used_xl_functions"], cell_code_lines + alias_lines
        )
        runtime_import_block = self._format_from_runtime_import(internals_import_names)
        internals_lines: list[str] = ["from __future__ import annotations", ""]
        used_readers = sorted(self._used_readers)
        if self._internals_needs_datetime_import(cell_code_lines + alias_lines):
            internals_lines.extend(["import datetime", ""])
        if used_readers:
            internals_lines.append(self._format_from_module_import("_readers", used_readers))
        if runtime_import_block:
            internals_lines.append(runtime_import_block)
        if used_readers or runtime_import_block:
            internals_lines.append("")
        internals_lines.append("# --- Formula cell functions ---\n")
        internals_lines.extend(cell_code_lines)
        internals_lines.extend(alias_lines)
        internals_lines.extend(
            self._emit_resolver_lines(parts["blank_rects"] if parts["blank_rects"] else None)
        )
        internals_py = "\n".join(internals_lines).rstrip() + "\n"

        series_setter_names: list[str] = []
        series_reader_names: list[str] = []
        series_compute_names: list[str] = []
        series_range_reader_names: list[str] = []
        reader_leaves: dict[str, dict[str, object]] | None = None
        reader_ranges: dict[str, dict[str, object]] | None = None
        api_helpers_py: str | None = None
        readers_py: str | None = None
        output_leaves_py: str | None = None
        reader_lines: list[str] = []
        output_leaf_lines: list[str] = []
        setter_lines: list[str] = []
        leaf_index_imports: list[str] = []
        output_leaves_imports: list[str] = []
        helper_imports: list[str] = []
        public_reader_imports: list[str] = []
        output_helper_imports: list[str] = []
        if series_bindings is not None:
            if bindings_workbook is None:
                raise ValueError("bindings_workbook is required when series_bindings is set")
            # Route coercion helpers to `_api_helpers`, input leaf maps / readers to
            # `_readers`, and output leaf tables to `_output_leaves` so `api.py`
            # stays focused on the public surface.
            emit_input = self._series_bindings_have_input(series_bindings)
            emit_readers = self._series_bindings_have_readers(series_bindings)
            emit_output = self._series_bindings_have_output(series_bindings)
            export_with_aliases = self._export_addresses_with_aliases(
                _all_cells,
                series_public_addresses,
            )
            from excel_grapher.series_bindings.output_helper_index import (
                build_output_helper_index,
                output_helper_names,
            )

            output_helper_index = None
            if emit_output:
                output_helper_index = build_output_helper_index(
                    cast("DependencyGraph", self._public_graph()),
                    series_bindings,
                    workbook=bindings_workbook,
                    export_addresses=export_with_aliases,
                    address_helpers=address_helpers,
                )
                output_helper_imports = output_helper_names(output_helper_index)
                from excel_grapher.series_bindings.compute_codegen import emit_output_leaves_block

                output_leaf_lines = emit_output_leaves_block(
                    cast("DependencyGraph", self._public_graph()),
                    bindings_workbook,
                    series_bindings,
                    export_addresses=export_with_aliases,
                )
                output_leaves_py = self._emit_output_leaves_module(output_leaf_lines)
                output_leaves_imports = self._series_output_leaves_imports(output_leaf_lines)
            if emit_readers:
                from excel_grapher.series_bindings.setter_codegen import emit_readers_block

                reader_lines = emit_readers_block(
                    cast("DependencyGraph", self._public_graph()),
                    bindings_workbook,
                    series_bindings,
                    export_addresses=export_with_aliases,
                    series_docstring_callback=series_docstring_callback,
                    docstring_renderer=docstring_renderer,
                )
                readers_py = self._emit_readers_module(reader_lines)
                leaf_index_imports = self._series_reader_leaf_index_imports(reader_lines)
                public_reader_imports = self._series_reader_public_imports(reader_lines)
            setter_lines = self._emit_series_binding_setters(
                series_bindings,
                bindings_workbook,
                export_addresses=_all_cells,
                public_addresses=series_public_addresses,
                include_helpers=not emit_input,
                include_readers=not emit_readers,
                include_leaf_indexes=not emit_readers,
                include_leaves_tables=not emit_output,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
                helper_index=output_helper_index,
                address_helpers=address_helpers,
            )
            if emit_input:
                api_helpers_py = self._emit_api_helpers_module()
                helper_imports = self._series_helper_imports(setter_lines)
            (
                series_setter_names,
                series_reader_names,
                series_compute_names,
            ) = self._series_binding_public_names(series_bindings)
            series_range_reader_names = self._series_binding_emitted_range_reader_names(
                reader_lines if reader_lines else setter_lines
            )
            if self._reader_index is not None:
                from excel_grapher.series_bindings.reader_index import (
                    reader_index_as_discovery_dicts,
                )

                reader_leaves, reader_ranges = reader_index_as_discovery_dicts(self._reader_index)
            else:
                reader_leaves, reader_ranges = self._series_binding_reader_discovery(
                    cast("DependencyGraph", self._public_graph()),
                    series_bindings,
                    workbook=bindings_workbook,
                    export_addresses=_all_cells,
                )

        groups_manifest = self._series_binding_groups_manifest(series_bindings)
        discovery_lines = self._emit_series_binding_discovery_lines(
            series_setter_names,
            series_compute_names,
            groups_manifest,
            reader_names=series_reader_names,
            reader_leaves=reader_leaves,
            reader_ranges=reader_ranges,
        )
        coverage_export_addresses = (
            self._export_addresses_with_aliases(_all_cells, series_public_addresses)
            if series_bindings is not None
            else _all_cells
        )
        emit_compute_all = self._should_emit_compute_all(
            normalized_targets,
            series_bindings=series_bindings,
            bindings_workbook=bindings_workbook,
            export_addresses=coverage_export_addresses,
            include_compute_all=include_compute_all,
        )
        compute_all_lines = (
            self._emit_compute_all_block(normalized_targets) if emit_compute_all else []
        )
        api_body_lines: list[str] = [
            "def make_context(inputs: dict[str, object] | None = None) -> EvalContext:",
            '    """Create an EvalContext with merged inputs."""',
            "    merged = prepare_context_inputs(DEFAULT_INPUTS, CONSTANTS, inputs)",
            (
                "    return EvalContext("
                "inputs=merged, resolver=_resolve_formula, "
                f"iterative_enabled={bool(self._iterate_enabled)}, "
                f"iterate_count={int(self._iterate_count)}, "
                f"iterate_delta={float(self._iterate_delta)!r})"
            ),
            "",
            "",
            *setter_lines,
            *([] if not setter_lines else [""]),
            *discovery_lines,
            *(["", ""] if compute_all_lines else []),
            *compute_all_lines,
        ]
        api_body_text = "\n".join(api_body_lines)
        runtime_entry_names = ["EvalContext", "prepare_context_inputs"]
        # TARGETS may reference handlers by name (`xl_cell`, `xl_range_rows`) without a call.
        if re.search(r"\bxl_cell\b", api_body_text):
            runtime_entry_names.append("xl_cell")
        if re.search(r"\bxl_range\b", api_body_text):
            runtime_entry_names.append("xl_range")
        if re.search(r"\bxl_range_rows\b", api_body_text):
            runtime_entry_names.append("xl_range_rows")
        if emit_compute_all and self._iterate_enabled:
            runtime_entry_names.append("xl_iterative_compute")
        if re.search(r"\bXlErrorException\b", api_body_text):
            runtime_entry_names.append("XlErrorException")
        runtime_entry_names = sorted(set(runtime_entry_names))
        runtime_imports = self._format_from_runtime_import(runtime_entry_names)

        api_import_lines: list[str] = [
            "from __future__ import annotations",
            "",
        ]
        if "warnings." in api_body_text:
            api_import_lines.extend(["import warnings", ""])
        if helper_imports:
            api_import_lines.append(self._format_from_module_import("_api_helpers", helper_imports))
        readers_api_imports = sorted({*leaf_index_imports, *public_reader_imports})
        if readers_api_imports:
            # Leaf indexes are used by setters; public readers are re-exported via `__all__`.
            api_import_lines.append(
                self._format_from_module_import("_readers", readers_api_imports)
            )
        if output_leaves_imports:
            api_import_lines.append(
                self._format_from_module_import("_output_leaves", output_leaves_imports)
            )
        internals_api_imports = ["_resolve_formula", *output_helper_imports]
        api_import_lines.extend(
            [
                "from .data import CONSTANTS, DEFAULT_INPUTS",
                self._format_from_module_import("internals", internals_api_imports),
                runtime_imports,
                "",
                "",
            ]
        )

        # Public surface on `api.py`: discovery → setters → readers → computes.
        # Readers are defined in `_readers.py` and re-exported here.
        api_exports = [
            *(["compute_all"] if emit_compute_all else []),
            "make_context",
            "list_setters",
            "list_readers",
            "list_computes",
        ]
        if reader_leaves is not None:
            api_exports.append("list_reader_leaves")
        if reader_ranges is not None:
            api_exports.append("list_reader_ranges")
        if groups_manifest is not None:
            api_exports.append("list_groups")
        api_exports.extend(series_setter_names)
        api_exports.extend(series_reader_names)
        api_exports.extend(series_range_reader_names)
        api_exports.extend(series_compute_names)
        api_py = "\n".join(
            [
                *api_import_lines,
                *api_body_lines,
                f"__all__ = {api_exports!r}",
                "",
            ]
        )

        # Package `__all__` mirrors `api` plus `DEFAULT_INPUTS`.
        all_exports = [*api_exports, "DEFAULT_INPUTS"]
        init_lines = [
            "from __future__ import annotations",
            "",
        ]
        # Import names are sorted for isort; `__all__` keeps the deliberate public order.
        init_lines.append(self._format_from_module_import("api", sorted(api_exports), noqa="F401"))
        init_lines.extend(
            [
                "from .data import DEFAULT_INPUTS  # noqa: F401",
                "",
                f"__all__ = {all_exports!r}",
                "",
            ]
        )
        init_py = "\n".join(init_lines)

        modules = {
            "__init__.py": init_py,
            "api.py": api_py,
            "data.py": data_py,
            "runtime.py": runtime_py,
            "internals.py": internals_py,
        }
        if api_helpers_py is not None:
            modules["_api_helpers.py"] = api_helpers_py
        if readers_py is not None:
            modules["_readers.py"] = readers_py
        if output_leaves_py is not None:
            modules["_output_leaves.py"] = output_leaves_py
        return modules

    def _workbook_sort_addresses(self, addresses: Iterable[str]) -> list[str]:
        """Return addresses sorted by workbook sheet order, then row, then column."""
        materialized = [normalize_address(addr) for addr in addresses]
        if not materialized:
            return []
        sheet_order = getattr(self.graph, "sheet_order", None)
        if sheet_order:
            return sort_node_keys(materialized, sheet_order=sheet_order)
        return sorted(materialized)

    def _generate_parts(
        self,
        targets: list[str],
        *,
        dependency_targets: list[str] | None = None,
        constant_types: set[str] | None = None,
        constant_ranges: Sequence[str] | None = None,
        constant_blanks: bool = False,
        input_ranges: Sequence[str] | None = None,
        blank_ranges: Sequence[str] | None = None,
        series_bindings: WorkbookSeriesBindings | None = None,
        bindings_workbook: Path | str | None = None,
    ) -> GenerationParts:
        """Generate shared intermediate artifacts for single-file and modular exports."""
        self._reset_transient_state()

        blank_rects = normalize_blank_range_specs(blank_ranges)

        normalized_targets = [normalize_address(t) for t in targets]
        normalized_dependency_targets = (
            [normalize_address(t) for t in dependency_targets]
            if dependency_targets is not None
            else normalized_targets
        )

        # Collect all dependencies for all targets.
        #
        # When we are given a real excel_grapher.DependencyGraph, prefer its dependency edges
        # as the single source of truth for the export surface area. This ensures the exported
        # package can evaluate the full excel_grapher dependency closure (for the workbook's
        # cached state) without missing-cell KeyErrors.
        all_cells = self._collect_all_cells(normalized_dependency_targets)

        # Generate formula cell functions and collect used xl_* functions
        self._emitted.clear()
        cell_code_lines: list[str] = []
        used_xl_functions: set[str] = set()
        formula_cells: set[str] = set()
        formula_emit_order: list[str] = []

        def _track_cell(address: str) -> None:
            if address in self._emitted:
                return
            self._emitted.add(address)
            node = self.graph.get_node(address)
            if node is not None and _node_has_formula(node):
                normalized = normalize_address(address)
                formula_emit_order.append(normalized)
                self._temp_var_counter = 0
                ast = self._get_or_parse_ast(address)
                assert ast is not None
                prev_cell = self._formula_cell_address
                self._formula_cell_address = normalized
                try:
                    self._emit_ast(ast)
                finally:
                    self._formula_cell_address = prev_cell

        for address in all_cells:
            _track_cell(address)

        # If dynamic OFFSET was used, ensure the runtime implementation is embedded.
        #
        # For graphs without excel_grapher dependency edges (common in unit tests that only add
        # nodes), we expand the export surface area to include the full graph so dynamic OFFSET
        # reads do not hit missing-cell KeyErrors. For real workbook graphs (excel_grapher
        # closure), we intentionally *do not* widen the export surface area here.
        if self._needs_offset_runtime:
            used_xl_functions.add("xl_offset")
            if self._needs_index_ref_runtime:
                used_xl_functions.add("xl_index_ref")
            if not self._used_graph_closure:
                all_graph_cells = list(self.graph.leaf_keys()) + list(self.graph.formula_keys())
                if self._offset_runtime_sheets:
                    all_graph_cells = [
                        addr
                        for addr in all_graph_cells
                        if parse_address(normalize_address(addr))[0] in self._offset_runtime_sheets
                    ]
                for address in all_graph_cells:
                    _track_cell(address)
                    all_cells.append(address)

        # Build the reader index after the export surface is final (including any
        # OFFSET widening) so discovery and body rewrite share one map. Clear
        # `_used_readers` first: the probe `_emit_ast` pass above may have
        # touched the (previously unset) index without contributing to emit.
        self._used_readers.clear()
        if (
            series_bindings is not None
            and bindings_workbook is not None
            and self._series_bindings_have_readers(series_bindings)
        ):
            from excel_grapher.series_bindings.reader_index import build_reader_index

            series_public = self._series_binding_public_addresses(
                series_bindings,
                bindings_workbook,
            )
            self._reader_index = build_reader_index(
                cast("DependencyGraph", self._public_graph()),
                series_bindings,
                workbook=bindings_workbook,
                export_addresses=self._export_addresses_with_aliases(
                    all_cells,
                    series_public,
                ),
            )
        else:
            self._reader_index = None

        self._plan_shape_helpers(formula_emit_order)
        helper_lines = self._emit_shape_helpers()
        if helper_lines:
            cell_code_lines.extend(helper_lines)

        for address in self._workbook_sort_addresses(formula_emit_order):
            formula_cells.add(address)
            cell_code_lines.append(self._emit_cell(address))
            cell_code_lines.append("")
            cell_code_lines.append("")

            ast = self._get_or_parse_ast(address)
            assert ast is not None
            self._note_operators_fastpath_from_ast(ast)
            used_xl_functions.update(self._extract_xl_functions(ast))

        self._assert_raise_only_cell_boundary(cell_code_lines)

        # Always include per-call evaluation scaffolding.
        # XlError is commonly referenced by generated code (error literals, IF/IFERROR, and
        # potentially leaf inputs), so keep it available.
        runtime_symbols = set(used_xl_functions) | {
            "EvalContext",
            "coerce_inputs_dict",
            "prepare_context_inputs",
            "xl_cell",
            "xl_eval",
            "xl_raise",
            "XlError",
            "XlErrorException",
        }
        # Multi-cell targets materialize through the eager range boundary handler.
        if any(
            handler == "xl_range_rows"
            for _, handler in self._targets_to_entries(normalized_targets)
        ):
            runtime_symbols.add("xl_range_rows")
        # Binding-aligned `read_*_range` helpers call `xl_range` even when no formula
        # AST or target-entry coalescing would otherwise pull it into the runtime.
        if series_bindings is not None and self._series_bindings_may_emit_range_readers(
            series_bindings
        ):
            runtime_symbols.add("xl_range")
        if series_bindings is not None and self._series_bindings_have_readers(series_bindings):
            # Generated `read_*` annotations return `CellValue`.
            runtime_symbols.add("CellValue")
        if self._iterate_enabled:
            runtime_symbols.add("xl_iterative_compute")
        include_dep_tracking = self._include_dep_tracking(series_bindings)
        runtime_code = emit_runtime(
            runtime_symbols,
            include_offset_table=False,
            include_dep_tracking=include_dep_tracking,
            include_operators_fastpath=self._needs_operators_fastpath,
        )
        runtime_code = runtime_code.rstrip()

        normalized_constant_types = self._normalize_constant_types(constant_types)
        normalized_constant_ranges = self._normalize_constant_ranges(constant_ranges)
        normalized_input_ranges = self._normalize_input_ranges(input_ranges)
        include_constants = bool(constant_types or constant_ranges or constant_blanks)
        graph_classification = None
        if not include_constants:
            graph_classification = self._get_graph_leaf_classification()
        inputs_block_lines, constants_block_lines = self._emit_default_inputs_lines(
            all_cells,
            constant_types=normalized_constant_types,
            constant_ranges=normalized_constant_ranges,
            constant_blanks=constant_blanks,
            graph_classification=graph_classification,
            include_constants=include_constants,
            input_ranges=normalized_input_ranges,
        )
        if graph_classification is not None and constants_block_lines:
            include_constants = True

        return {
            "runtime_code": runtime_code,
            "inputs_block_lines": inputs_block_lines,
            "constants_block_lines": constants_block_lines,
            "cell_code_lines": cell_code_lines,
            "formula_cells": self._workbook_sort_addresses(formula_cells),
            "all_cells": all_cells,
            "needs_offset_table": self._needs_offset_runtime,
            "targets": normalized_targets,
            "has_constants": include_constants,
            "used_xl_functions": frozenset(used_xl_functions),
            "blank_rects": blank_rects,
        }

    def _resolve_targets(self, targets: Sequence[str] | None) -> list[str]:
        if targets is not None:
            return self._expand_target_tokens(targets)
        target_keys = getattr(self._public_graph(), "target_keys", None)
        inferred_targets: list[str] = []
        if callable(target_keys):
            inferred_targets = list(target_keys())
        if not inferred_targets:
            raise ValueError(
                "No export targets were provided and the graph has no target-marked nodes."
            )
        return [normalize_address(t) for t in inferred_targets]

    def _collect_all_cells(self, targets: list[str]) -> list[str]:
        """Collect an ordered list of addresses to emit for the given targets.

        For excel_grapher.DependencyGraph instances, this uses the graph dependency edges and
        evaluation order as the export closure. For other GraphLike implementations, it falls
        back to the CodeGenerator AST-based dependency walk.
        """
        projected_targets = [self._map_address_to_projected(t) for t in targets]
        # Prefer graph-driven closure when excel_grapher provides an evaluation order AND
        # has dependency edges populated. Many unit tests build a DependencyGraph with nodes
        # only (no edges); for those we must fall back to AST-based dependency discovery.
        eval_order = getattr(self.graph, "evaluation_order", None)
        if callable(eval_order):
            # Heuristic: only use graph edges if any target has at least one dependency edge.
            # (Graphs constructed via create_dependency_graph(...) will satisfy this for
            # non-leaf targets; test graphs that only add nodes will not.)
            has_edges = any(bool(self.graph.get_dependencies(t)) for t in projected_targets)
            if not has_edges:
                return self._collect_all_cells_via_ast(projected_targets)

            self._used_graph_closure = True
            closure: set[str] = set()
            stack = list(projected_targets)
            while stack:
                addr = normalize_address(stack.pop())
                if addr in closure:
                    continue
                node = self.graph.get_node(addr)
                if node is None:
                    continue
                closure.add(addr)
                for dep in self.graph.get_dependencies(addr):
                    dep_n = normalize_address(dep)
                    owner = dep_n
                    resolve = getattr(self.graph, "resolve_endpoint", None)
                    if callable(resolve):
                        resolved = resolve(dep_n)
                        if resolved is not None:
                            owner = resolved
                    owner_n = normalize_address(owner)
                    if owner_n not in closure and self.graph.get_node(owner_n) is not None:
                        stack.append(owner_n)

            if self._iterate_enabled:
                ordered = []
            else:
                try:
                    ordered = list(eval_order(strict=False))
                except CycleError:
                    ordered = []
                except Exception:
                    ordered = []

            out: list[str] = []
            seen: set[str] = set()
            for addr in ordered:
                a = normalize_address(addr)
                if a in seen or a not in closure:
                    continue
                seen.add(a)
                out.append(a)
            for addr in self._workbook_sort_addresses(closure):
                if addr not in seen:
                    out.append(addr)
            return out

        return self._collect_all_cells_via_ast(projected_targets)

    def _collect_all_cells_via_ast(self, targets: list[str]) -> list[str]:
        """AST-based dependency walk (works for GraphLike test doubles)."""
        out: list[str] = []
        seen: set[str] = set()
        for target in targets:
            deps = self._collect_dependencies(target)
            for dep in deps:
                dep_n = normalize_address(dep)
                if dep_n not in seen:
                    seen.add(dep_n)
                    out.append(dep_n)
        return out

    def _emit_default_inputs_lines(
        self,
        all_cells: list[str],
        *,
        constant_types: set[str],
        constant_ranges: list[tuple[str, int, int, int, int]],
        constant_blanks: bool,
        graph_classification: dict[str, str] | None,
        include_constants: bool,
        input_ranges: list[tuple[str, int, int, int, int]],
    ) -> tuple[list[str], list[str]]:
        needed_leaves = self._collect_needed_leaves(all_cells)

        if include_constants:
            inputs, constants = self._classify_leaf_nodes(
                needed_leaves,
                constant_types=constant_types,
                constant_ranges=constant_ranges,
                constant_blanks=constant_blanks,
                input_ranges=input_ranges,
            )
        else:
            inputs, constants = self._classification_from_graph(graph_classification, needed_leaves)
            inputs, constants = self._apply_input_ranges_override(
                needed_leaves, constants, input_ranges
            )

        default_lines = [
            "# --- Default inputs (leaf cells) ---",
            *self._emit_leaf_store_lines("DEFAULT_INPUTS", inputs),
        ]
        constants_lines: list[str] = []
        if constants:
            constants_lines = [
                "# --- Constant leaf values ---",
                *self._emit_leaf_store_lines("CONSTANTS", constants, frozen=True),
            ]
        return default_lines, constants_lines

    def _emit_leaf_store_lines(
        self,
        name: str,
        addresses: set[str] | list[str],
        *,
        frozen: bool = False,
    ) -> list[str]:
        """Emit a nested `sheet -> {(row, col): value}` mapping literal.

        When `frozen` is True, wrap the outer store and each sheet map in
        `MappingProxyType` so generated `CONSTANTS` fail closed on mutation.
        """
        by_sheet: dict[str, list[tuple[int, int, str]]] = {}
        for key in addresses:
            node = self.graph.get_node(key)
            value = 0 if node is None else node.value
            sheet, row, col = self._leaf_coords(key, node)
            by_sheet.setdefault(sheet, []).append((row, col, self._py_literal(value)))

        if not by_sheet:
            if frozen:
                return [f"{name} = MappingProxyType({{}})"]
            return [f"{name} = {{}}"]

        outer_open = "MappingProxyType({" if frozen else "{"
        inner_open = "MappingProxyType({" if frozen else "{"
        inner_close = "})," if frozen else "},"
        outer_close = "})" if frozen else "}"

        lines = [f"{name} = {outer_open}"]
        for sheet in sorted(by_sheet):
            lines.append(f"    {sheet!r}: {inner_open}")
            for row, col, lit in sorted(by_sheet[sheet], key=lambda item: (item[0], item[1])):
                lines.append(f"        ({row}, {col}): {lit},")
            lines.append(f"    {inner_close}")
        lines.append(outer_close)
        return lines

    def _leaf_coords(self, key: str, node: object | None) -> tuple[str, int, int]:
        """Resolve a leaf address to `(sheet, row, col)`."""
        if node is not None:
            sheet = getattr(node, "sheet", None)
            row = getattr(node, "row", None)
            column_index = getattr(node, "column_index", None)
            if isinstance(sheet, str) and isinstance(row, int) and isinstance(column_index, int):
                return sheet, row, column_index
            column = getattr(node, "column", None)
            if isinstance(sheet, str) and isinstance(row, int) and isinstance(column, str):
                return sheet, row, int(fastpyxl.utils.cell.column_index_from_string(column))
        return parse_cell_coords(key)

    def _emit_offset_cell_table_lines(self, all_cells: list[str]) -> list[str]:
        lines: list[str] = []
        lines.append("# --- Cell lookup table for dynamic OFFSET ---")
        lines.append("_CELL_TABLE = {")
        for address in all_cells:
            normalized = normalize_address(address)
            sheet, cell = parse_address(normalized)
            col_str, row = fastpyxl.utils.cell.coordinate_from_string(cell)
            col = fastpyxl.utils.cell.column_index_from_string(col_str)
            lines.append(f"    ({repr(sheet)}, {row}, {col}): {repr(normalized)},")
        lines.append("}")
        return lines
