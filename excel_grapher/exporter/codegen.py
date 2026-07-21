"""Code generator for converting Excel formulas to Python code."""

from __future__ import annotations

import ast as py_ast
import json
import re
from collections.abc import Iterable, Mapping, Sequence, Set
from pathlib import Path
from typing import TYPE_CHECKING, Any, Protocol, TypedDict, cast

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import (
    format_range_key,
    parse_address,
    quote_sheet_if_needed,
    sort_node_keys,
)
from excel_grapher.core.address_keys import (
    normalize_key as normalize_address,
)
from excel_grapher.core.formula_ast import (
    AddressHoleNode,
    AddressLeafKind,
)
from excel_grapher.core.operators_fastpath import MIN_OPERATOR_FASTPATH_CELLS
from excel_grapher.evaluator.errors import FormulaGroupKeyError, MissingNormalizedFormulaError
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
from excel_grapher.exporter.projection import ProjectedAddress
from excel_grapher.grapher.blank_ranges import BlankRangeRect, normalize_blank_range_specs
from excel_grapher.grapher.formula_groups import collect_holes, serialize_address_leaf
from excel_grapher.grapher.graph import CycleError
from excel_grapher.grapher.node import NodeKind, locate_cell
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


# Comparison operators emitted via xl_compare / xl_map_compare.
_COMPARE_OPS = frozenset({"=", "<>", "<", ">", "<=", ">="})

# Binary operators emitted as native Python with coercion helpers.
_ARITHMETIC_OPS = frozenset({"+", "-", "*", "/", "^"})

# Functions whose single argument is emitted as a lazily-evaluated thunk so the
# exported runtime can catch raised Excel errors. Mirrors the evaluator's
# AST-level special cases; other IS functions propagate argument errors there.
_THUNK_ARG_FUNCTIONS = frozenset({"ISERROR", "ISNA", "ISBLANK"})


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
        self._offset_runtime_sheets: set[str] = set()
        self._temp_var_counter = 0  # Counter for unique temp variable names
        self._ast_cache: dict[str, AstNode] = {}
        self._used_graph_closure: bool = False
        self._formula_cell_address: str | None = None
        # When emitting a `_group_*` helper, no-arg ROW/COLUMN use this param name.
        self._group_member_param: str | None = None
        self._hole_param_by_slot: dict[int, str] | None = None
        # member address -> ProjectedAddress with group key + binding params
        self._group_member_exports: dict[str, ProjectedAddress] = {}

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
        self._group_member_param = None
        self._hole_param_by_slot = None
        self._group_member_exports.clear()

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

    def _map_address_to_projected(self, address: str) -> ProjectedAddress:
        """Map a public address through projection, then formula-group occupancy."""
        normalized = normalize_address(address)
        manifest = self._projection_manifest()
        if manifest is None:
            projected = ProjectedAddress(address=normalized, parameters=None)
        else:
            projected = manifest.map_to_projected(normalized)

        # Overlay formula-group ownership (hand-built or coalesced).
        # GraphLike test doubles may lack occupancy / iteration; skip overlay then.
        locate_graph = self._public_graph()
        try:
            location = locate_cell(cast(Any, locate_graph), normalized)
        except (TypeError, ValueError):
            location = None
        if location is None or location.kind is NodeKind.cell:
            return projected

        owner = locate_graph.get_node(location.node_key)
        if owner is None:
            return ProjectedAddress(address=location.node_key, parameters=None)
        skeleton = getattr(owner, "skeleton", None)
        member_bindings = getattr(owner, "member_bindings", None)
        if skeleton is None or member_bindings is None:
            return ProjectedAddress(address=location.node_key, parameters=None)

        bindings = member_bindings.get(normalized)
        if bindings is None:
            return ProjectedAddress(address=location.node_key, parameters=None)

        return ProjectedAddress(
            address=location.node_key,
            parameters={
                "member": normalized,
                "bindings": tuple(serialize_address_leaf(b) for b in bindings),
            },
        )

    def _projection_alias_map(
        self,
        public_addresses: Iterable[str],
        export_addresses: Iterable[str],
    ) -> dict[str, str]:
        exported = frozenset(normalize_address(addr) for addr in export_addresses)
        aliases: dict[str, str] = {}
        for address in public_addresses:
            public_addr = normalize_address(address)
            projected = self._map_address_to_projected(public_addr)
            projected_addr = normalize_address(projected.address)
            # Formula-group members get dedicated wrappers; skip alias map.
            if projected.parameters is not None:
                continue
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
        for raw in targets:
            self._reject_formula_group_key_target(raw)
        named_ranges, named_range_ranges = self._named_range_maps()
        roots = expand_targets_to_roots(
            targets,
            sheetnames=self._graph_sheetnames(targets=targets),
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
        return [normalize_address(format_key(sheet, a1)) for sheet, a1 in roots]

    def _reject_formula_group_key_target(self, address: str) -> None:
        """Reject multi-cell formula-group keys as export targets.

        Public export targets must be member cell addresses (or ordinary ranges of
        cell nodes). A `RangeKey` / `UnionKey` that owns a formula-group template
        is not a supported target — matching `FormulaEvaluator`.

        Bare defined-name tokens are left for `expand_targets_to_roots`.
        """
        try:
            normalized = normalize_address(address)
        except ValueError:
            return
        node = self.graph.get_node(normalized)
        if node is not None and getattr(node, "skeleton", None) is not None:
            raise FormulaGroupKeyError(normalized)

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
            # Error literals raise in the exported error channel.
            return f"xl_raise(XlError.{node.error.name})"

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

        if isinstance(node, AddressHoleNode):
            if self._hole_param_by_slot is None:
                raise ValueError("AddressHoleNode outside formula-group helper emission")
            param = self._hole_param_by_slot.get(node.slot)
            if param is None:
                raise ValueError(f"Missing binding parameter for hole slot {node.slot}")
            if node.kind is AddressLeafKind.cell:
                return f"xl_cell(ctx, {param})"
            if node.kind is AddressLeafKind.range:
                return f"xl_range(ctx, {param})"
            if node.kind is AddressLeafKind.whole_column:
                return f"xl_range(ctx, {param})"
            if node.kind is AddressLeafKind.whole_row:
                return f"xl_range(ctx, {param})"
            raise ValueError(f"Unsupported address hole kind: {node.kind!r}")

        raise ValueError(f"Unknown AST node type: {type(node)}")

    def _emit_range(self, node: RangeNode) -> str:
        """Emit a range as a lazy `Range` value resolved through the context.

        For A1:B3, emits: xl_range(ctx, "S!A1:B3"). Consumers evaluate cells
        positionally; unused cells are never evaluated.
        """
        return self._emit_range_address(node.start, node.end)

    def _emit_range_address(self, start: str, end: str) -> str:
        """Emit an xl_range call for a normalized start/end address pair."""
        sheet, r1, c1, r2, c2 = self._range_coords(start, end)
        start_cell = f"{fastpyxl.utils.cell.get_column_letter(c1)}{r1}"
        end_cell = f"{fastpyxl.utils.cell.get_column_letter(c2)}{r2}"
        return f"xl_range(ctx, {repr(format_range_key(sheet, start_cell, end_cell))})"

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
        include_helpers: bool = True,
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
            include_helpers=include_helpers,
            series_docstring_callback=series_docstring_callback,
            docstring_renderer=docstring_renderer,
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
    ) -> tuple[list[str], list[str]]:
        """Return declared public setter and compute function names.

        Without groups the names sort alphabetically (flat export); with
        view-level groups they follow the grouped export order.
        """
        from excel_grapher.series_bindings.groups import (
            bindings_have_groups,
            grouped_public_names,
        )
        from excel_grapher.series_bindings.workflow import compute_names, setter_names

        if bindings_have_groups(bindings):
            return grouped_public_names(bindings)
        return setter_names(bindings), compute_names(bindings)

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
    def _emit_series_binding_discovery_lines(
        setter_names: Sequence[str],
        compute_names: Sequence[str],
        groups_manifest: Mapping[str, Any] | None = None,
    ) -> list[str]:
        """Emit generated-code helpers that list public series-binding functions."""
        lines = [
            "def list_setters() -> list[str]:",
            '    """Return generated series-binding setter function names."""',
            f"    return {list(setter_names)!r}",
            "",
            "",
            "def list_computes() -> list[str]:",
            '    """Return generated series-binding compute function names."""',
            f"    return {list(compute_names)!r}",
        ]
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
        names.update({"xl_cell", "xl_eval"})
        if "xl_range(" in blob:
            names.add("xl_range")
        if "xl_raise(" in blob:
            names.add("xl_raise")
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
        left = self._emit_ast(node.left)
        right = self._emit_ast(node.right)
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
        operand = self._emit_ast(node.operand)
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
            arg_expr = self._emit_ast(node.args[0])
            return f"{func_name}(lambda: ({arg_expr}))"

        emitted_args = [self._emit_ast(arg) for arg in node.args]
        args = ", ".join(emitted_args)
        return f"{func_name}({args})"

    def _next_temp_var(self) -> str:
        """Generate a unique temporary variable name."""
        self._temp_var_counter += 1
        return f"_t{self._temp_var_counter}"

    def _emit_lazy_error_fallback(self, node: FunctionCallNode, name: str) -> str:
        """Emit IFERROR/IFNA as thunked runtime calls with try/except semantics.

        IFERROR(value, value_if_error) evaluates the fallback only when
        evaluating ``value`` produces any Excel error (raised
        ``XlErrorException`` or ``XlError`` sentinel). IFNA does so only for
        ``#N/A`` and re-raises other errors.
        """
        if len(node.args) < 2:
            return "xl_raise(XlError.VALUE)"

        value_expr = self._emit_ast(node.args[0])
        fallback_expr = self._emit_ast(node.args[1])

        if name == "IFERROR":
            func = "xl_iferror"
        elif name == "IFNA":
            func = "xl_ifna"
        else:
            raise ValueError(f"Unsupported lazy error fallback function: {name!r}")

        return f"{func}(lambda: ({value_expr}), lambda: ({fallback_expr}))"

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

        cond_expr = self._emit_ast(node.args[0])
        true_expr = self._emit_ast(node.args[1])
        false_expr = self._emit_ast(node.args[2]) if len(node.args) > 2 else "False"

        # Excel-style boolean coercion is not Python truthiness:
        # - "FALSE" should behave like False
        # - "0" should produce #VALUE! (per to_bool)
        # `xl_bool` keeps lazy branch evaluation while raising coercion errors.
        bool_var = self._next_temp_var()
        return f"(({true_expr}) if ({bool_var} := xl_bool({cond_expr})) else ({false_expr}))"

    def _emit_row(self, node: FunctionCallNode) -> str:
        if not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)):
            if self._group_member_param is not None:
                return f"xl_formula_row({self._group_member_param})"
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

        return f"xl_row({self._emit_ast(arg)})"

    def _emit_column(self, node: FunctionCallNode) -> str:
        if not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)):
            if self._group_member_param is not None:
                return f"xl_formula_column({self._group_member_param})"
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

        return f"xl_column({self._emit_ast(arg)})"

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

        return f"xl_columns({self._emit_ast(arg)})"

    def _emit_choose(self, node: FunctionCallNode) -> str:
        """Emit CHOOSE as chained conditionals for lazy evaluation.

        CHOOSE(index_num, value1, [value2], ...)

        Emits as chained conditionals that only evaluate the selected value.
        This is critical for breaking circular references that Excel handles
        via lazy evaluation.
        """
        if len(node.args) < 2:
            return "xl_raise(XlError.VALUE)"

        index_expr = self._emit_ast(node.args[0])
        value_exprs = [self._emit_ast(arg) for arg in node.args[1:]]

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

    def _ast_needs_array_operator_branch(self, node: AstNode) -> bool:
        """Return whether a subtree can *evaluate to* a range/array at runtime.

        Only array-producing nodes require the operator broadcast branch: ranges,
        multi-cell `OFFSET`, non-scalar `INDEX` slices, and the pass-through
        functions (`IF`/`IFERROR`/`IFNA`/`CHOOSE`) when a returned branch is
        itself an array. Scalar-returning functions (e.g. `SUM`, `MATCH`,
        `VLOOKUP`) never yield arrays even when their arguments contain ranges,
        so they take the inlined scalar path without a guard.
        """
        if isinstance(node, RangeNode):
            return True
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
            return "xl_raise(XlError.REF)"

        rows_expr = self._emit_ast(rows_node)
        cols_expr = self._emit_ast(cols_node)
        height_expr = "None" if height_node is None else self._emit_ast(height_node)
        width_expr = "None" if width_node is None else self._emit_ast(width_node)

        return f"xl_offset(ctx, {ref_info}, {rows_expr}, {cols_expr}, {height_expr}, {width_expr})"

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
            return "xl_raise(XlError.REF)"

        rows_expr = self._emit_ast(rows_node)
        cols_expr = self._emit_ast(cols_node)
        height_expr = "None" if height_node is None else self._emit_ast(height_node)
        width_expr = "None" if width_node is None else self._emit_ast(width_node)

        return f"xl_offset_ref({ref_info}, {rows_expr}, {cols_expr}, {height_expr}, {width_expr})"

    def _group_helper_name(self, group_key: str) -> str:
        base = address_to_python_name(normalize_address(group_key))
        if base.startswith("cell_"):
            base = base[len("cell_") :]
        return f"_group_{base}"

    def _emit_group_helper(self, group_key: str) -> str:
        """Emit one parameterized `_group_*` helper for a formula-group node."""
        normalized = normalize_address(group_key)
        node = self.graph.get_node(normalized)
        skeleton = None if node is None else getattr(node, "skeleton", None)
        if skeleton is None:
            raise ValueError(f"Not a formula-group node: {normalized}")

        holes = collect_holes(skeleton)
        helper = self._group_helper_name(normalized)
        params = ", ".join(["ctx", "member", *[f"b{i}" for i in range(len(holes))]])
        lines = [f"def {helper}({params}):"]
        doc = f"Formula group: {normalized}".replace("'''", "\\'''")
        lines.append(f"    '''{doc}.'''")
        self._temp_var_counter = 0
        prev_holes = self._hole_param_by_slot
        self._hole_param_by_slot = {hole.slot: f"b{i}" for i, hole in enumerate(holes)}
        prev_cell = self._formula_cell_address
        prev_member = self._group_member_param
        self._formula_cell_address = normalized
        self._group_member_param = "member"
        try:
            expr = self._emit_ast(skeleton)
        finally:
            self._hole_param_by_slot = prev_holes
            self._formula_cell_address = prev_cell
            self._group_member_param = prev_member
        lines.append(f"    return {expr}")
        return "\n".join(lines)

    def _emit_group_member_wrapper_lines(self) -> list[str]:
        """Emit thin `cell_*` wrappers that call `_group_*` with member bindings."""
        if not self._group_member_exports:
            return []
        lines = ["# --- Formula-group member wrappers ---", ""]
        for member, projected in sorted(self._group_member_exports.items()):
            params = projected.parameters or {}
            bindings = params.get("bindings") or ()
            helper = self._group_helper_name(projected.address)
            wrapper = address_to_python_name(member)
            call_args = ", ".join([repr(member), *[repr(b) for b in bindings]])
            call = f"{helper}(ctx, {call_args})"
            lines.extend(
                [
                    f"def {wrapper}(ctx):",
                    f"    return {call}",
                    "",
                ]
            )
        return lines

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
        """
        funcs: set[str] = set()

        if isinstance(node, ErrorNode):
            # Error literals raise via xl_raise and reference the XlError enum
            funcs.add("XlError")
            funcs.add("xl_raise")
        elif isinstance(node, RangeNode):
            # Ranges emit lazy xl_range(ctx, ...) calls
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
                if not node.args or (
                    len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)
                ):
                    # Group helpers resolve no-arg ROW via xl_formula_row(member).
                    funcs.add("xl_formula_row")
                else:
                    funcs.add("xl_row")
                    ref = node.args[0]
                    if isinstance(ref, FunctionCallNode) and ref.name.upper() == "OFFSET":
                        funcs.add("xl_offset_ref")
                        for off_arg in ref.args:
                            funcs.update(self._extract_xl_functions(off_arg))
                    else:
                        funcs.update(self._extract_xl_functions(ref))
            elif upper_name == "COLUMN":
                if not node.args or (
                    len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)
                ):
                    funcs.add("xl_formula_column")
                else:
                    funcs.add("xl_column")
                    ref = node.args[0]
                    if isinstance(ref, FunctionCallNode) and ref.name.upper() == "OFFSET":
                        funcs.add("xl_offset_ref")
                        for off_arg in ref.args:
                            funcs.update(self._extract_xl_functions(off_arg))
                    else:
                        funcs.update(self._extract_xl_functions(ref))
            elif upper_name == "COLUMNS":
                funcs.add("xl_columns")
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
        elif isinstance(node, AddressHoleNode):
            if node.kind is AddressLeafKind.cell:
                funcs.add("xl_cell")
            else:
                funcs.add("xl_range")

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
        for addr in public_addresses:
            projected = self._map_address_to_projected(addr)
            if projected.parameters is not None:
                member = str(projected.parameters.get("member") or normalize_address(addr))
                self._group_member_exports[member] = projected
        lines.extend(self._emit_group_member_wrapper_lines())
        alias_lines = self._emit_projection_alias_lines(_all_cells, public_addresses)
        lines.extend(alias_lines)
        lines.extend(
            self._emit_resolver_lines(parts["blank_rects"] if parts["blank_rects"] else None)
        )

        # Generate entry point helpers
        lines.append("def make_context(inputs: dict[str, object] | None = None) -> EvalContext:")
        lines.append('    """Create an EvalContext with merged inputs."""')
        lines.append("    merged: dict[str, object] = dict(DEFAULT_INPUTS)")
        if parts["has_constants"]:
            lines.append("    merged.update(CONSTANTS)")
        lines.append("    if inputs is not None:")
        lines.append("        merged.update(inputs)")
        lines.append(
            "    return EvalContext("
            "inputs=coerce_inputs_dict(merged), resolver=_resolve_formula, "
            f"iterative_enabled={bool(self._iterate_enabled)}, "
            f"iterate_count={int(self._iterate_count)}, "
            f"iterate_delta={float(self._iterate_delta)!r})"
        )
        lines.append("")
        lines.append("")
        series_setter_names: list[str] = []
        series_compute_names: list[str] = []
        if series_bindings is not None:
            if bindings_workbook is None:
                raise ValueError("bindings_workbook is required when series_bindings is set")
            series_setter_names, series_compute_names = self._series_binding_public_names(
                series_bindings
            )
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
        lines.extend(
            self._emit_series_binding_discovery_lines(
                series_setter_names,
                series_compute_names,
                self._series_binding_groups_manifest(series_bindings),
            )
        )
        lines.append("")
        lines.append("")
        lines.append("TARGETS = {")
        for target, handler in self._targets_to_entries(normalized_targets):
            lines.append(f"    {repr(target)}: {handler},")
        lines.append("}")
        lines.append("")
        lines.append("")
        lines.append(
            "def compute_all(ctx: EvalContext | None = None, *, "
            "inputs: dict[str, object] | None = None) -> dict[str, object]:"
        )
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
        needs_range_helper = any(handler == "xl_range_rows" for _, handler in targets_entries)

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
        for addr in public_addresses:
            projected = self._map_address_to_projected(addr)
            if projected.parameters is not None:
                member = str(projected.parameters.get("member") or normalize_address(addr))
                self._group_member_exports[member] = projected
        group_wrapper_lines = self._emit_group_member_wrapper_lines()
        internals_import_names = self._internals_runtime_import_names(
            parts["used_xl_functions"], cell_code_lines + group_wrapper_lines + alias_lines
        )
        runtime_import_block = self._format_from_runtime_import(internals_import_names)
        internals_lines: list[str] = ["from __future__ import annotations", ""]
        if runtime_import_block:
            internals_lines.append(runtime_import_block)
            internals_lines.append("")
        internals_lines.append("# --- Formula cell functions ---\n")
        internals_lines.extend(cell_code_lines)
        internals_lines.extend(group_wrapper_lines)
        internals_lines.extend(alias_lines)
        internals_lines.extend(
            self._emit_resolver_lines(parts["blank_rects"] if parts["blank_rects"] else None)
        )
        internals_py = "\n".join(internals_lines).rstrip() + "\n"

        runtime_entry_names = ["EvalContext", "coerce_inputs_dict", "xl_cell"]
        if needs_range_helper:
            runtime_entry_names.append("xl_range_rows")
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
            "",
            "def make_context(inputs: dict[str, object] | None = None) -> EvalContext:",
            '    """Create an EvalContext with merged inputs."""',
            "    merged: dict[str, object] = dict(DEFAULT_INPUTS)",
            "    merged.update(CONSTANTS)",
            "    if inputs is not None:",
            "        merged.update(inputs)",
            (
                "    return EvalContext("
                "inputs=coerce_inputs_dict(merged), resolver=_resolve_formula, "
                f"iterative_enabled={bool(self._iterate_enabled)}, "
                f"iterate_count={int(self._iterate_count)}, "
                f"iterate_delta={float(self._iterate_delta)!r})"
            ),
            "",
            "",
        ]
        series_setter_names: list[str] = []
        series_compute_names: list[str] = []
        api_helpers_py: str | None = None
        if series_bindings is not None:
            if bindings_workbook is None:
                raise ValueError("bindings_workbook is required when series_bindings is set")
            # Route the verbose input-coercion helpers to a private `_api_helpers`
            # module so `api.py` stays focused on the public surface.
            emit_input = self._series_bindings_have_input(series_bindings)
            setter_lines = self._emit_series_binding_setters(
                series_bindings,
                bindings_workbook,
                export_addresses=_all_cells,
                public_addresses=series_public_addresses,
                include_helpers=not emit_input,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
            )
            if emit_input:
                api_helpers_py = self._emit_api_helpers_module()
                helper_imports = self._series_helper_imports(setter_lines)
                if helper_imports:
                    api_lines.insert(4, f"from ._api_helpers import {', '.join(helper_imports)}")
            api_lines.extend(setter_lines)
            series_setter_names, series_compute_names = self._series_binding_public_names(
                series_bindings
            )
            api_lines.append("")
        groups_manifest = self._series_binding_groups_manifest(series_bindings)
        api_lines.extend(
            self._emit_series_binding_discovery_lines(
                series_setter_names,
                series_compute_names,
                groups_manifest,
            )
        )
        api_lines.append("")
        api_lines.append("")
        api_lines.append("TARGETS = {")
        for target, handler in targets_entries:
            api_lines.append(f"    {repr(target)}: {handler},")
        api_lines.extend(
            [
                "}",
                "",
                "",
                "def compute_all(ctx: EvalContext | None = None, *, "
                "inputs: dict[str, object] | None = None) -> dict[str, object]:",
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

        api_exports = ["compute_all", "make_context", "list_setters", "list_computes"]
        if groups_manifest is not None:
            api_exports.append("list_groups")
        api_exports.extend(series_setter_names)
        api_exports.extend(series_compute_names)
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

        modules = {
            "__init__.py": init_py,
            "api.py": api_py,
            "data.py": data_py,
            "runtime.py": runtime_py,
            "internals.py": internals_py,
        }
        if api_helpers_py is not None:
            modules["_api_helpers.py"] = api_helpers_py
        if groups_manifest is not None:
            modules["groups.json"] = json.dumps(groups_manifest, indent=2) + "\n"
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
            if node is not None and getattr(node, "skeleton", None) is not None:
                normalized = normalize_address(address)
                formula_emit_order.append(normalized)
                return
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
            node = self.graph.get_node(address)
            skeleton = None if node is None else getattr(node, "skeleton", None)
            if skeleton is not None:
                cell_code_lines.append(self._emit_group_helper(address))
                cell_code_lines.append("")
                cell_code_lines.append("")
                prev_holes = self._hole_param_by_slot
                holes = collect_holes(skeleton)
                self._hole_param_by_slot = {h.slot: f"b{i}" for i, h in enumerate(holes)}
                try:
                    self._note_operators_fastpath_from_ast(skeleton)
                    used_xl_functions.update(self._extract_xl_functions(skeleton))
                finally:
                    self._hole_param_by_slot = prev_holes
                continue

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
        projected_targets = []
        for t in targets:
            projected = self._map_address_to_projected(t)
            projected_targets.append(projected.address)
            if projected.parameters is not None:
                member = str(projected.parameters.get("member") or normalize_address(t))
                self._group_member_exports[member] = projected
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
            node = self.graph.get_node(normalize_address(target))
            if node is not None and getattr(node, "skeleton", None) is not None:
                # Formula-group: use graph edges when present; otherwise the group key alone.
                deps = [normalize_address(target)]
                for dep in self.graph.get_dependencies(normalize_address(target)):
                    deps.append(normalize_address(dep))
            else:
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
