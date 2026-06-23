"""Code generator for converting Excel formulas to Python code."""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence, Set
from pathlib import Path
from typing import TYPE_CHECKING, Any, Protocol, TypedDict, cast

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    normalize_key as normalize_address,
)
from excel_grapher.core.address_keys import (
    parse_address,
    quote_sheet_if_needed,
    sort_node_keys,
)
from excel_grapher.core.array_results import spill_footprint_addresses
from excel_grapher.core.excel_function_meta import numpy_array_arg_indices
from excel_grapher.core.operators_fastpath import MIN_OPERATOR_FASTPATH_CELLS
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
from excel_grapher.grapher.graph import CycleError
from excel_grapher.grapher.parser import format_key
from excel_grapher.grapher.target_expansion import (
    expand_targets_to_roots,
    split_range_target_on_colon,
)

__all__ = ["CodeGenerator", "GenerationParts", "GraphLike", "GraphNode"]

if TYPE_CHECKING:
    from excel_grapher.exporter.projection import ProjectionManifest
    from excel_grapher.grapher import DependencyGraph  # noqa: F401
    from excel_grapher.series_bindings.docstring_renderers import SeriesDocstringRendererSpec
    from excel_grapher.series_bindings.docstrings import SeriesBindingDocstringCallbackSpec
    from excel_grapher.series_bindings.types import InputSeries, WorkbookSeriesBindings


class GraphNode(Protocol):
    formula: str | None
    normalized_formula: str | None
    value: object | None


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


# Operators that need wrapper functions for Excel semantics (error propagation)
_BINARY_OPS = {
    "+": "xl_add",
    "-": "xl_sub",
    "*": "xl_mul",
    "/": "xl_div",
    "^": "xl_pow",
    "=": "xl_eq",
    "<>": "xl_ne",
    "<": "xl_lt",
    ">": "xl_gt",
    "<=": "xl_le",
    ">=": "xl_ge",
}

# Unary operators that need wrapper functions
_UNARY_OPS = {
    "-": "xl_neg",
    "+": "xl_pos",
    "%": "xl_percent",
}


class CodeGenerator:
    """Generates Python code from Excel formulas."""

    def __init__(
        self,
        graph: DependencyGraph | GraphLike,
        *,
        iterate_enabled: bool | None = None,
        iterate_count: int = 100,
        iterate_delta: float = 0.001,
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
        """
        self.graph = graph
        self._iterate_enabled = iterate_enabled
        self._iterate_count = iterate_count
        self._iterate_delta = iterate_delta
        self._emitted: set[str] = set()
        self._needs_offset_runtime = False  # Set to True if dynamic OFFSET is used
        self._needs_index_ref_runtime = False  # OFFSET(INDEX(...), ...) requires xl_index_ref
        self._needs_operators_fastpath = False  # Large array binary ops / SUMPRODUCT
        self._needs_array_results = False  # Top-level arrays, spill reads, spill blocking
        self._offset_runtime_sheets: set[str] = set()
        self._temp_var_counter = 0  # Counter for unique temp variable names
        self._ast_cache: dict[str, AstNode] = {}
        self._used_graph_closure: bool = False
        self._formula_cell_address: str | None = None

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
        self._needs_array_results = False
        self._offset_runtime_sheets.clear()
        self._temp_var_counter = 0
        self._ast_cache.clear()
        self._used_graph_closure = False
        self._formula_cell_address = None

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
        from excel_grapher.series_bindings.ranges import expand_data_range_for_graph

        public_graph = cast("DependencyGraph", self._public_graph())
        addresses: set[str] = set()
        for series in bindings.get("series", []):
            if not isinstance(series, dict):
                continue
            data_range = series.get("data_range")
            if not isinstance(data_range, str):
                continue
            addresses.update(
                normalize_address(addr)
                for addr in expand_data_range_for_graph(
                    public_graph,
                    data_range,
                    workbook=workbook,
                )
            )
        return frozenset(addresses)

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
            replacement_node = self.graph.get_node(replacement)
            if replacement_node is not None and replacement_node.formula is not None:
                replacement_fn = address_to_python_name(replacement)
                lines.extend(
                    [
                        f"def {public_fn}(ctx):",
                        f"    return xl_eval(ctx, {repr(replacement)}, {replacement_fn})",
                        "",
                    ]
                )
            else:
                lines.extend(
                    [
                        f"def {public_fn}(ctx):",
                        f"    return xl_cell(ctx, {repr(replacement)})",
                        "",
                    ]
                )
        return lines

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
        if node is None or node.formula is None:
            return None

        nf = node.normalized_formula
        if nf is None or not isinstance(nf, str) or not nf.strip():
            raise MissingNormalizedFormulaError(normalized)
        ast = parse(nf.strip())
        self._ast_cache[normalized] = ast
        return ast

    def _emit_ast(self, node: AstNode) -> str:
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
            return f"XlError.{node.error.name}"

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
        """Emit a range as a 2D nested list of cell evaluations.

        The outer list contains rows, inner lists contain columns.
        For A1:B3, emits: [[xl_eval(ctx, "S!A1", cell_s_a1), xl_eval(ctx, "S!B1", ...)], ...]
        """
        rows = self._range_addresses_2d(node.start, node.end)
        row_strs = []
        for row_addrs in rows:
            cell_calls = [self._emit_cell_eval_for_range(addr) for addr in row_addrs]
            row_strs.append("[" + ", ".join(cell_calls) + "]")
        # Model ranges as object-dtype ndarrays so they fit `CellValue` and work
        # with runtime helpers like `flatten(*args)`.
        return f"np.array([{', '.join(row_strs)}], dtype=object)"

    def _emit_cell_eval_for_range(self, address: str) -> str:
        """Emit a range member read, with spill projection only when needed (#284)."""
        normalized = normalize_address(address)
        if self.graph is None:
            return f"xl_cell(ctx, {repr(normalized)})"
        node = self.graph.get_node(normalized)
        if node is not None and node.formula is not None:
            func_name = address_to_python_name(normalized)
            return (
                f"scalar_for_range_member({repr(normalized)}, "
                f"xl_eval(ctx, {repr(normalized)}, {func_name}))"
            )
        if node is not None:
            return f"xl_cell(ctx, {repr(normalized)})"
        return f"xl_cell_in_range(ctx, {repr(normalized)})"

    def _emit_cell_eval(self, address: str) -> str:
        normalized = normalize_address(address)
        if self.graph is None:
            return f"xl_cell(ctx, {repr(normalized)})"
        node = self.graph.get_node(normalized)
        if node is not None and node.formula is not None:
            func_name = address_to_python_name(normalized)
            return f"xl_eval(ctx, {repr(normalized)}, {func_name})"
        return f"xl_cell(ctx, {repr(normalized)})"

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
        series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
        docstring_renderer: SeriesDocstringRendererSpec = "google",
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
            series_docstring_callback=series_docstring_callback,
            docstring_renderer=docstring_renderer,
        )

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
    def _emitted_function_names(lines: Sequence[str]) -> list[str]:
        names: list[str] = []
        for line in lines:
            if not line.startswith("def "):
                continue
            name, _open_paren, _rest = line[4:].partition("(")
            if name:
                names.append(name)
        return names

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
                    start_addr = self._format_cell_address(sheet, row, start)
                    end_addr = self._format_cell_address(sheet, row, prev)
                    if start == prev:
                        row_entries.append((start_addr, "xl_cell"))
                    else:
                        row_entries.append((f"{start_addr}:{end_addr}", "xl_range"))
                    start = prev = col
                start_addr = self._format_cell_address(sheet, row, start)
                end_addr = self._format_cell_address(sheet, row, prev)
                if start == prev:
                    row_entries.append((start_addr, "xl_cell"))
                else:
                    row_entries.append((f"{start_addr}:{end_addr}", "xl_range"))

            col_entries: list[tuple[str, str]] = []
            for col, rows in col_groups.items():
                rows = sorted(rows)
                start = prev = rows[0]
                for row in rows[1:]:
                    if row == prev + 1:
                        prev = row
                        continue
                    start_addr = self._format_cell_address(sheet, start, col)
                    end_addr = self._format_cell_address(sheet, prev, col)
                    if start == prev:
                        col_entries.append((start_addr, "xl_cell"))
                    else:
                        col_entries.append((f"{start_addr}:{end_addr}", "xl_range"))
                    start = prev = row
                start_addr = self._format_cell_address(sheet, start, col)
                end_addr = self._format_cell_address(sheet, prev, col)
                if start == prev:
                    col_entries.append((start_addr, "xl_cell"))
                else:
                    col_entries.append((f"{start_addr}:{end_addr}", "xl_range"))

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
    def _broadcast_array_shapes(
        left: tuple[int, int] | None,
        right: tuple[int, int] | None,
    ) -> tuple[int, int] | None:
        """Broadcast two static array shapes for top-level binary operators."""
        if left is None:
            return right
        if right is None:
            return left
        if left == (1, 1):
            return right
        if right == (1, 1):
            return left
        if left == right:
            return left
        return None

    def _ast_result_shape(self, node: AstNode) -> tuple[int, int] | None:
        """Infer a static result shape from an AST subtree when possible."""
        if isinstance(node, RangeNode):
            rows = self._range_addresses_2d(node.start, node.end)
            if not rows or not rows[0]:
                return None
            return (len(rows), len(rows[0]))
        if isinstance(node, CellRefNode):
            return (1, 1)
        if isinstance(node, (NumberNode, StringNode, BoolNode, ErrorNode, EmptyArgNode)):
            return (1, 1)
        if isinstance(node, BinaryOpNode):
            return self._broadcast_array_shapes(
                self._ast_result_shape(node.left),
                self._ast_result_shape(node.right),
            )
        if isinstance(node, UnaryOpNode):
            return self._ast_result_shape(node.operand)
        if isinstance(node, FunctionCallNode):
            return None
        return None

    def _infer_top_level_array_shape(self, ast: AstNode) -> tuple[int, int] | None:
        """Return spill footprint shape for formulas that yield multi-cell arrays."""
        shape = self._ast_result_shape(ast)
        if shape is None or shape == (1, 1):
            return None
        return shape

    def _spill_footprint_slots_for_formula(self, anchor: str) -> set[str]:
        """Return spill footprint addresses for a top-level array formula."""
        ast = self._get_or_parse_ast(anchor)
        if ast is None:
            return set()
        shape = self._infer_top_level_array_shape(ast)
        if shape is None:
            return set()
        return set(spill_footprint_addresses(anchor, shape))

    def _spill_occupied_addresses(self, closure: frozenset[str] | None = None) -> frozenset[str]:
        """Occupied spill slots that can block array formulas in the export closure."""
        if self.graph is None or closure is None:
            return frozenset()
        occupied: set[str] = set()
        for addr in closure:
            node = self.graph.get_node(addr)
            if node is None or node.formula is None:
                continue
            for slot in self._spill_footprint_slots_for_formula(addr):
                slot_node = self.graph.get_node(slot)
                if slot_node is None:
                    continue
                if slot_node.formula is not None or slot_node.value is not None:
                    occupied.add(slot)
        return frozenset(occupied)

    def _emit_spill_occupancy_lines(self, closure: frozenset[str] | None = None) -> list[str]:
        """Emit spill occupancy helper used by ``EvalContext``."""
        occupied = sorted(self._spill_occupied_addresses(closure))
        if not occupied:
            return [
                "def _spill_is_occupied(_address: str) -> bool:",
                "    return False",
                "",
            ]
        lines = ["_SPILL_OCCUPIED = frozenset({"]
        lines.extend(f"    {address!r}," for address in occupied)
        lines.extend(
            [
                "})",
                "",
                "def _spill_is_occupied(address: str) -> bool:",
                "    return address in _SPILL_OCCUPIED",
                "",
            ]
        )
        return lines

    def _eval_context_ctor_kwargs(self) -> str:
        """Keyword arguments for generated ``EvalContext(...)`` calls."""
        parts = [
            "inputs=coerce_inputs_dict(merged)",
            "resolver=_resolve_formula",
        ]
        if self._needs_array_results:
            parts.append("spill_is_occupied=_spill_is_occupied")
        parts.extend(
            [
                f"iterative_enabled={bool(self._iterate_enabled)}",
                f"iterate_count={int(self._iterate_count)}",
                f"iterate_delta={float(self._iterate_delta)!r}",
            ]
        )
        return ", ".join(parts)

    @staticmethod
    def _internals_runtime_import_names(
        used_xl_functions: Set[str], cell_code_lines: list[str]
    ) -> list[str]:
        """Names from the embedded runtime that formula cell bodies reference as globals."""
        blob = "\n".join(cell_code_lines)
        if "def " not in blob:
            return []
        names = set(used_xl_functions)
        names.discard("numpy")
        names.update({"xl_cell", "xl_eval"})
        if "xl_cell_in_range" in blob:
            names.add("xl_cell_in_range")
        if "scalar_for_range_member" in blob:
            names.add("scalar_for_range_member")
        if "numpy" in used_xl_functions or "np." in blob or "np.array" in blob:
            names.add("np")
        if "XlError" in blob:
            names.add("XlError")
        if "ExcelRange(" in blob:
            names.add("ExcelRange")
        return sorted(names)

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
            if node is None or node.formula is not None:
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

    def _emit_binary_op(self, node: BinaryOpNode) -> str:
        """Emit a binary operation."""
        left = self._emit_ast(node.left)
        right = self._emit_ast(node.right)
        op = node.op

        # Concatenation: & -> xl_concat
        if op == "&":
            return f"xl_concat({left}, {right})"

        # All other operators use wrapper functions for error propagation
        if op in _BINARY_OPS:
            func = _BINARY_OPS[op]
            return f"{func}({left}, {right})"

        raise ValueError(f"Unknown operator: {op}")

    def _emit_unary_op(self, node: UnaryOpNode) -> str:
        """Emit a unary operation."""
        operand = self._emit_ast(node.operand)
        op = node.op
        if op in _UNARY_OPS:
            func = _UNARY_OPS[op]
            return f"{func}({operand})"
        raise ValueError(f"Unknown unary operator: {op}")

    def _emit_function_call(self, node: FunctionCallNode) -> str:
        """Emit a function call.

        For functions that need numpy arrays (LOOKUP, VLOOKUP, HLOOKUP, INDEX,
        MATCH, SUMPRODUCT), range arguments are wrapped with np.array().
        IF, OFFSET are handled specially.
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

        # Functions that need numpy arrays for their array/table arguments
        needs_numpy_wrap = numpy_array_arg_indices(upper_name)

        emitted_args = []
        for i, arg in enumerate(node.args):
            emitted = self._emit_ast(arg)
            # Wrap range arguments with np.array() for functions that need it
            # Use dtype=object to preserve original Python types (mixed str/int/float)
            if i in needs_numpy_wrap and isinstance(arg, RangeNode):
                emitted = f"np.array({emitted}, dtype=object)"
            emitted_args.append(emitted)

        args = ", ".join(emitted_args)
        return f"{func_name}({args})"

    def _next_temp_var(self) -> str:
        """Generate a unique temporary variable name."""
        self._temp_var_counter += 1
        return f"_t{self._temp_var_counter}"

    def _emit_lazy_error_fallback(self, node: FunctionCallNode, name: str) -> str:
        """Emit IFERROR/IFNA as Python conditionals for lazy fallback evaluation.

        IFERROR(value, value_if_error) evaluates the fallback only when ``value``
        is any ``XlError``. IFNA(value, value_if_na) does so only for ``#N/A``.
        """
        if len(node.args) < 2:
            return "XlError.VALUE"

        value_expr = self._emit_ast(node.args[0])
        fallback_expr = self._emit_ast(node.args[1])
        var = self._next_temp_var()

        if name == "IFERROR":
            condition = f"isinstance(({var} := {value_expr}), XlError)"
        elif name == "IFNA":
            condition = f"(({var} := {value_expr}) == XlError.NA)"
        else:
            raise ValueError(f"Unsupported lazy error fallback function: {name!r}")

        return f"(({fallback_expr}) if {condition} else {var})"

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
            return "XlError.VALUE"

        cond_expr = self._emit_ast(node.args[0])
        true_expr = self._emit_ast(node.args[1])
        false_expr = self._emit_ast(node.args[2]) if len(node.args) > 2 else "False"

        # Excel-style boolean coercion is not Python truthiness:
        # - "FALSE" should behave like False
        # - "0" should produce #VALUE! (per to_bool)
        # We must coerce via to_bool(), and keep lazy branch evaluation.
        cond_var = self._next_temp_var()
        bool_var = self._next_temp_var()
        return (
            f"({bool_var} if isinstance(({bool_var} := to_bool(({cond_var} := {cond_expr}))), XlError) "
            f"else (({true_expr}) if {bool_var} else ({false_expr})))"
        )

    def _emit_row(self, node: FunctionCallNode) -> str:
        if not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)):
            addr = self._formula_cell_address
            if addr is None:
                return "XlError.VALUE"
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

        return f"xl_row({self._emit_ast(arg)})"

    def _emit_column(self, node: FunctionCallNode) -> str:
        if not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)):
            addr = self._formula_cell_address
            if addr is None:
                return "XlError.VALUE"
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

        return f"xl_column({self._emit_ast(arg)})"

    def _emit_columns(self, node: FunctionCallNode) -> str:
        if len(node.args) < 1:
            return "XlError.VALUE"

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

        return f"xl_columns({self._emit_ast(arg)})"

    def _emit_choose(self, node: FunctionCallNode) -> str:
        """Emit CHOOSE as chained conditionals for lazy evaluation.

        CHOOSE(index_num, value1, [value2], ...)

        Emits as chained conditionals that only evaluate the selected value.
        This is critical for breaking circular references that Excel handles
        via lazy evaluation.
        """
        if len(node.args) < 2:
            return "XlError.VALUE"

        index_expr = self._emit_ast(node.args[0])
        value_exprs = [self._emit_ast(arg) for arg in node.args[1:]]

        # Store index in temp vars to avoid evaluating twice and to keep typing clean.
        # We coerce via to_int() (Excel-style numeric coercion + error propagation)
        # to avoid `int(CellValue)` in generated code (which type-checkers reject).
        var = self._next_temp_var()
        idx_var = self._next_temp_var()

        # Build chained conditionals: if idx==1 then val1 else if idx==2 then val2 ...
        # Start from the innermost (last value or VALUE error for out of bounds)
        result = "XlError.VALUE"
        for i, val_expr in reversed(list(enumerate(value_exprs, start=1))):
            result = f"(({val_expr}) if {idx_var} == {i} else ({result}))"

        # Wrap with error/bounds checking
        return (
            f"({var} if isinstance(({var} := {index_expr}), XlError) "
            f"else ({idx_var} if isinstance(({idx_var} := to_int({var})), XlError) "
            f"else XlError.VALUE if {idx_var} < 1 or {idx_var} > {len(value_exprs)} else {result}))"
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

    def _emit_offset(self, node: FunctionCallNode) -> str:
        """Emit OFFSET function, trying static resolution first.

        OFFSET(reference, rows, cols, [height], [width])

        If all offset arguments are constants, resolves to direct cell/range reference.
        Otherwise, falls back to runtime xl_offset() function.
        """
        if len(node.args) < 3:
            # Invalid OFFSET - need at least reference, rows, cols
            return "XlError.VALUE"

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
            return nrows == 1
        if col_omitted:
            return ncols == 1
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
            else self._emit_ast(node.args[1])
        )
        col_expr = (
            "None"
            if len(node.args) < 3 or isinstance(node.args[2], EmptyArgNode)
            else self._emit_ast(node.args[2])
        )
        self._needs_offset_runtime = True
        self._needs_index_ref_runtime = True
        return f"xl_offset(ctx, xl_index_ref({base_ref_info}, {row_expr}, {col_expr}), 0.0, 0.0)"

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

    def _note_operators_fastpath_from_ast(self, node: AstNode) -> None:
        if self._max_array_extent_in_ast(node) >= MIN_OPERATOR_FASTPATH_CELLS:
            self._needs_operators_fastpath = True

    def _note_array_results_from_ast(self, node: AstNode) -> None:
        if self._infer_top_level_array_shape(node) is not None:
            self._needs_array_results = True

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
            return "XlError.REF"

        target_col_str = fastpyxl.utils.cell.get_column_letter(target_col)

        if height == 1 and width == 1:
            # Single cell reference
            target_addr = f"{quote_sheet_if_needed(base_sheet)}!{target_col_str}{target_row}"
            return self._emit_cell_eval(target_addr)
        else:
            # Range reference - emit as 2D array
            end_row = target_row + height - 1
            end_col = target_col + width - 1
            end_col_str = fastpyxl.utils.cell.get_column_letter(end_col)

            start_addr = f"{quote_sheet_if_needed(base_sheet)}!{target_col_str}{target_row}"
            end_addr = f"{quote_sheet_if_needed(base_sheet)}!{end_col_str}{end_row}"

            # Generate 2D array like _emit_range does
            rows_list = self._range_addresses_2d(start_addr, end_addr)
            row_strs = []
            for row_addrs in rows_list:
                cell_calls = [self._emit_cell_eval(addr) for addr in row_addrs]
                row_strs.append("[" + ", ".join(cell_calls) + "]")
            return f"np.array([{', '.join(row_strs)}], dtype=object)"

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
                return "XlError.VALUE"
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
                return "XlError.REF"

            row_expr = (
                "None"
                if len(ref_node.args) < 2 or isinstance(ref_node.args[1], EmptyArgNode)
                else self._emit_ast(ref_node.args[1])
            )
            col_expr = (
                "None"
                if len(ref_node.args) < 3 or isinstance(ref_node.args[2], EmptyArgNode)
                else self._emit_ast(ref_node.args[2])
            )
            self._needs_index_ref_runtime = True
            ref_info = f"xl_index_ref({base_ref_info}, {row_expr}, {col_expr})"
        else:
            # If reference is not a simple cell, we can't handle it
            return "XlError.REF"

        rows_expr = self._emit_ast(rows_node)
        cols_expr = self._emit_ast(cols_node)
        height_expr = "None" if height_node is None else self._emit_ast(height_node)
        width_expr = "None" if width_node is None else self._emit_ast(width_node)

        return f"xl_offset(ctx, {ref_info}, {rows_expr}, {cols_expr}, {height_expr}, {width_expr})"

    def _emit_offset_ref(self, node: FunctionCallNode) -> str:
        if len(node.args) < 3:
            return "XlError.VALUE"

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
                return "XlError.VALUE"
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
                return "XlError.REF"

            row_expr = (
                "None"
                if len(ref_node.args) < 2 or isinstance(ref_node.args[1], EmptyArgNode)
                else self._emit_ast(ref_node.args[1])
            )
            col_expr = (
                "None"
                if len(ref_node.args) < 3 or isinstance(ref_node.args[2], EmptyArgNode)
                else self._emit_ast(ref_node.args[2])
            )
            self._needs_index_ref_runtime = True
            ref_info = f"xl_index_ref({base_ref_info}, {row_expr}, {col_expr})"
        else:
            return "XlError.REF"

        rows_expr = self._emit_ast(rows_node)
        cols_expr = self._emit_ast(cols_node)
        height_expr = "None" if height_node is None else self._emit_ast(height_node)
        width_expr = "None" if width_node is None else self._emit_ast(width_node)

        return f"xl_offset_ref({ref_info}, {rows_expr}, {cols_expr}, {height_expr}, {width_expr})"

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

        if node is None or node.formula is None:
            raise ValueError(f"Not a formula cell: {normalized}")

        lines: list[str] = []
        lines.append(f"def {func_name}(ctx):")
        doc = f"Formula: {node.formula}".replace("'''", "\\'''")
        if doc[-1] not in ".?!":
            doc = f"{doc}."
        lines.append(f"    '''{doc}'''")
        # Reset temp var counter for each cell to keep variable names short
        self._temp_var_counter = 0
        ast = self._get_or_parse_ast(normalized)
        assert ast is not None
        prev_cell = self._formula_cell_address
        self._formula_cell_address = normalized
        try:
            expr = self._emit_ast(ast)
        finally:
            self._formula_cell_address = prev_cell
        lines.append(f"    return {expr}")

        return "\n".join(lines)

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
            if node.formula is not None:
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
        - "numpy": numpy is needed for np.array() wrapping of ranges
        """
        funcs: set[str] = set()

        if isinstance(node, ErrorNode):
            # Error literal requires XlError enum
            funcs.add("XlError")
        elif isinstance(node, FunctionCallNode):
            upper_name = normalize_excel_function_name(node.name)

            # IF, IFERROR, CHOOSE are special - emitted as native Python conditionals
            if upper_name == "IF":
                funcs.add("XlError")
                funcs.add("to_bool")
            elif upper_name == "IFERROR" or upper_name == "IFNA":
                funcs.add("XlError")
            elif upper_name == "CHOOSE":
                funcs.add("XlError")
                funcs.add("to_int")
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

            # Check if this function needs numpy array wrapping for range args
            array_arg_indices = numpy_array_arg_indices(upper_name)
            if array_arg_indices:
                skip_index_array = (
                    upper_name == "INDEX"
                    and node.args
                    and isinstance(node.args[0], RangeNode)
                    and self._index_range_result_is_scalar(node.args[0], node)
                )
                for i, arg in enumerate(node.args):
                    if skip_index_array and upper_name == "INDEX":
                        break
                    if i in array_arg_indices and isinstance(arg, RangeNode):
                        funcs.add("numpy")
                        break
            for arg in node.args:
                funcs.update(self._extract_xl_functions(arg))
        elif isinstance(node, BinaryOpNode):
            # Binary operators use xl_* functions for error propagation
            if node.op == "&":
                funcs.add("xl_concat")
            elif node.op in _BINARY_OPS:
                funcs.add(_BINARY_OPS[node.op])
            funcs.update(self._extract_xl_functions(node.left))
            funcs.update(self._extract_xl_functions(node.right))
        elif isinstance(node, UnaryOpNode):
            # Unary operators use xl_* functions for error propagation
            if node.op in _UNARY_OPS:
                funcs.add(_UNARY_OPS[node.op])
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
    ) -> str:
        """Generate standalone Python code for target cells.

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
        lines: list[str] = [runtime_code, ""]

        lines.extend(parts["inputs_block_lines"])
        if parts["has_constants"]:
            lines.append("")
            lines.extend(parts["constants_block_lines"])
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
        lines.append("")
        if self._needs_array_results:
            lines.extend(self._emit_spill_occupancy_lines(frozenset(_all_cells)))

        # Generate entry point helpers
        lines.append("def make_context(inputs=None):")
        lines.append('    """Create an EvalContext with merged inputs."""')
        lines.append("    merged = dict(DEFAULT_INPUTS)")
        if parts["has_constants"]:
            lines.append("    merged.update(CONSTANTS)")
        lines.append("    if inputs is not None:")
        lines.append("        merged.update(inputs)")
        lines.append(f"    return EvalContext({self._eval_context_ctor_kwargs()})")
        lines.append("")
        lines.append("")
        if series_bindings is not None:
            if bindings_workbook is None:
                raise ValueError("bindings_workbook is required when series_bindings is set")
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
        lines.append("TARGETS = {")
        for target, handler in self._targets_to_entries(normalized_targets):
            lines.append(f"    {repr(target)}: {handler},")
        lines.append("}")
        lines.append("")
        lines.append("")
        lines.append("def compute_all(ctx=None, *, inputs=None):")
        lines.append('    """Compute all target cells and return results."""')
        lines.append("    if ctx is None:")
        lines.append("        ctx = make_context(inputs)")
        lines.append("    elif inputs is not None:")
        lines.append(
            "        warnings.warn("
            '"inputs will be ignored because ctx was provided", '
            "UserWarning, stacklevel=2)"
        )
        if self._iterate_enabled:
            lines.append("    return xl_iterative_compute(ctx, TARGETS)")
        else:
            lines.append(
                "    return {target: handler(ctx, target) for target, handler in TARGETS.items()}"
            )
        lines.append("")

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
    ) -> dict[str, str]:
        """Generate a multi-module Python package for target cells.

        Returns a mapping of module filenames to file contents. Callers choose how to
        lay out files on disk (e.g. a package directory).

        The generated package has five flat files:
        - __init__.py: exports compute_all and DEFAULT_INPUTS
        - api.py: compute_all, make_context, and series-binding set_* / compute_* functions
        - data.py: DEFAULT_INPUTS and CONSTANTS
        - runtime.py: embedded Excel runtime (emit_runtime)
        - internals.py: formula cell functions + resolver dispatch
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

        targets_entries = self._targets_to_entries(normalized_targets)
        needs_range_helper = any(handler == "xl_range" for _, handler in targets_entries)

        data_lines_out: list[str] = [
            "from __future__ import annotations",
            "",
            "# --- Default inputs (leaf cells) ---",
            *parts["inputs_block_lines"][1:],  # drop the comment already included above
            "",
        ]
        if parts["constants_block_lines"]:
            data_lines_out.extend(parts["constants_block_lines"])
        else:
            data_lines_out.append("CONSTANTS = {}")
        data_lines_out.append("")
        data_py = "\n".join(data_lines_out).rstrip() + "\n"

        runtime_py = runtime_code.rstrip() + "\n"

        alias_lines = self._emit_projection_alias_lines(_all_cells, public_addresses)
        internals_import_names = self._internals_runtime_import_names(
            parts["used_xl_functions"], cell_code_lines + alias_lines
        )
        runtime_import_block = self._format_from_runtime_import(internals_import_names)
        internals_lines: list[str] = ["from __future__ import annotations", ""]
        if runtime_import_block:
            internals_lines.append(runtime_import_block)
            internals_lines.append("")
        internals_lines.append("# --- Formula cell functions ---\n")
        internals_lines.extend(cell_code_lines)
        internals_lines.extend(alias_lines)
        internals_lines.extend(
            self._emit_resolver_lines(parts["blank_rects"] if parts["blank_rects"] else None)
        )
        internals_py = "\n".join(internals_lines).rstrip() + "\n"

        runtime_entry_names = ["EvalContext", "coerce_inputs_dict", "xl_cell"]
        if needs_range_helper:
            runtime_entry_names.append("xl_range")
        if self._iterate_enabled:
            runtime_entry_names.append("xl_iterative_compute")
        runtime_entry_names.sort()
        runtime_imports = self._format_from_runtime_import(runtime_entry_names)

        api_lines: list[str] = [
            "from __future__ import annotations",
            "",
            "from .data import CONSTANTS, DEFAULT_INPUTS",
            "from .internals import _resolve_formula",
            runtime_imports,
            "import warnings",
            "",
        ]
        if self._needs_array_results:
            api_lines.extend(self._emit_spill_occupancy_lines(frozenset(_all_cells)))
        api_lines.extend(
            [
                "",
                "def make_context(inputs=None):",
                '    """Create an EvalContext with merged inputs."""',
                "    merged = dict(DEFAULT_INPUTS)",
                "    merged.update(CONSTANTS)",
                "    if inputs is not None:",
                "        merged.update(inputs)",
                f"    return EvalContext({self._eval_context_ctor_kwargs()})",
                "",
                "",
            ]
        )
        series_setter_names: list[str] = []
        if series_bindings is not None:
            if bindings_workbook is None:
                raise ValueError("bindings_workbook is required when series_bindings is set")
            setter_lines = self._emit_series_binding_setters(
                series_bindings,
                bindings_workbook,
                export_addresses=_all_cells,
                public_addresses=series_public_addresses,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
            )
            api_lines.extend(setter_lines)
            series_setter_names = self._emitted_function_names(setter_lines)
            api_lines.append("")
        api_lines.append("TARGETS = {")
        for target, handler in targets_entries:
            api_lines.append(f"    {repr(target)}: {handler},")
        api_lines.extend(
            [
                "}",
                "",
                "",
                "def compute_all(ctx=None, *, inputs=None):",
                '    """Compute all target cells and return results."""',
                "    if ctx is None:",
                "        ctx = make_context(inputs)",
                "    elif inputs is not None:",
                "        warnings.warn(",
                '            "inputs will be ignored because ctx was provided",',
                "            UserWarning,",
                "            stacklevel=2,",
                "        )",
                (
                    "    return xl_iterative_compute(ctx, TARGETS)"
                    if self._iterate_enabled
                    else "    return {target: handler(ctx, target) for target, handler in TARGETS.items()}"
                ),
                "",
            ]
        )
        api_py = "\n".join(api_lines)

        api_exports = ["compute_all", "make_context"]
        api_exports.extend(series_setter_names)
        api_imports = ", ".join(api_exports)
        all_exports = api_exports + ["DEFAULT_INPUTS"]
        init_py = "\n".join(
            [
                "from __future__ import annotations",
                "",
                f"from .api import {api_imports}  # noqa: F401",
                "from .data import DEFAULT_INPUTS  # noqa: F401",
                "",
                f"__all__ = {all_exports!r}",
                "",
            ]
        )

        return {
            "__init__.py": init_py,
            "api.py": api_py,
            "data.py": data_py,
            "runtime.py": runtime_py,
            "internals.py": internals_py,
        }

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
            if node is not None and node.formula is not None:
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

        for address in self._workbook_sort_addresses(formula_emit_order):
            formula_cells.add(address)
            cell_code_lines.append(self._emit_cell(address))
            cell_code_lines.append("")
            cell_code_lines.append("")

            ast = self._get_or_parse_ast(address)
            assert ast is not None
            self._note_operators_fastpath_from_ast(ast)
            self._note_array_results_from_ast(ast)
            used_xl_functions.update(self._extract_xl_functions(ast))

        cell_blob = "\n".join(cell_code_lines)
        if "scalar_for_range_member" in cell_blob or "xl_cell_in_range" in cell_blob:
            self._needs_array_results = True
        if self._spill_occupied_addresses(frozenset(all_cells)):
            self._needs_array_results = True

        # Always include per-call evaluation scaffolding.
        # XlError is commonly referenced by generated code (error literals, IF/IFERROR, and
        # potentially leaf inputs), so keep it available.
        runtime_symbols = set(used_xl_functions) | {
            "EvalContext",
            "coerce_inputs_dict",
            "xl_cell",
            "xl_eval",
            "xl_range",
            "XlError",
        }
        if self._needs_array_results:
            runtime_symbols.add("xl_cell_in_range")
        if self._iterate_enabled:
            runtime_symbols.add("xl_iterative_compute")
        include_dep_tracking = self._include_dep_tracking(series_bindings)
        runtime_code = emit_runtime(
            runtime_symbols,
            include_offset_table=False,
            include_dep_tracking=include_dep_tracking,
            include_operators_fastpath=self._needs_operators_fastpath,
            include_array_results=self._needs_array_results,
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
                    if dep_n not in closure and self.graph.get_node(dep_n) is not None:
                        stack.append(dep_n)

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
        default_lines: list[str] = []
        default_lines.append("# --- Default inputs (leaf cells) ---")
        default_lines.append("DEFAULT_INPUTS = {")
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

        for key in sorted(inputs):
            node = self.graph.get_node(key)
            value = 0 if node is None else node.value
            default_lines.append(f"    {repr(key)}: {self._py_literal(value)},")
        default_lines.append("}")

        constants_lines: list[str] = []
        if constants:
            constants_lines.append("# --- Constant leaf values ---")
            constants_lines.append("CONSTANTS = {")
            for key in sorted(constants):
                node = self.graph.get_node(key)
                value = 0 if node is None else node.value
                constants_lines.append(f"    {repr(key)}: {self._py_literal(value)},")
            constants_lines.append("}")

        return default_lines, constants_lines

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
