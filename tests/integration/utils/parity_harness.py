"""Evaluator ↔ inverted-tree export helpers for function tests.

`evaluate_targets` still returns `FormulaEvaluator` results so existing
asserts keep Excel's sentinel error channel. It also emits a bindings-backed
inverted-tree package and checks those compute results against the evaluator
(#863 A1). Address-keyed `CodeGenerator.generate` was removed in #764.

Auto-bindings cover the demand cone of `formula_targets` plus constant series
for ranges those formulas read, not every formula on the graph.
"""

from __future__ import annotations

import uuid
from collections import defaultdict
from collections.abc import Sequence
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, cast

import pytest
from fastpyxl import load_workbook
from fastpyxl.utils.cell import get_column_letter

from excel_grapher import DependencyGraph, FormulaEvaluator
from excel_grapher.core.address_keys import (
    canonical_address,
    format_cell_key,
    format_range_key,
    parse_cell_coords,
)
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    FunctionCallNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    resolve_cell_ref,
)
from excel_grapher.core.types import XlError
from excel_grapher.exporter.inverted_tree.deps import (
    ast_literal_int,
    iter_range_addresses,
    iter_ref_addresses,
    normalize_excel_function_name,
    offset_index_destination,
    ref_window_corners,
    shift_range_corners,
)
from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.writeback import write_workbook
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    invoke_public_compute,
    load_package,
    series_entry,
    unload_package,
)

_LeafBlock = tuple[str, int, int, int, int]
_COMPARE_OPS = frozenset({"=", "<>", "<", ">", "<=", ">="})
_BOOL_FUNCS = frozenset(
    {
        "AND",
        "OR",
        "NOT",
        "TRUE",
        "FALSE",
        "ISBLANK",
        "ISERROR",
        "ISNA",
        "ISNUMBER",
        "ISTEXT",
    }
)
_STRING_FUNCS = frozenset({"TEXT", "T", "LEFT", "RIGHT", "MID", "CONCAT", "CONCATENATE"})
_EXPORT_ERROR_TYPES = frozenset({"XlError", "XlErrorException"})


def evaluate_targets(
    graph: DependencyGraph,
    targets: list[str],
    *,
    blank_ranges: tuple[str, ...] | list[str] | None = None,
) -> dict[str, object]:
    """Evaluate `targets` and assert inverted-tree export agrees.

    Returns:
        Evaluator results keyed by sheet-qualified address. Export errors are
        mapped to `XlError` sentinels before comparison so tests keep one
        channel.
    """
    with FormulaEvaluator(graph, blank_ranges=blank_ranges) as ev:
        expected = cast(dict[str, object], ev.evaluate(targets))
    exported = _export_targets(graph, targets, expected, blank_ranges=blank_ranges)
    for address in targets:
        want = expected[address]
        got = exported[address]
        assert _values_match(want, got), f"{address}: evaluator {want!r} != export {got!r}"
    return expected


def _export_targets(
    graph: DependencyGraph,
    targets: Sequence[str],
    expected: dict[str, object],
    *,
    blank_ranges: tuple[str, ...] | list[str] | None,
) -> dict[str, object]:
    formula_targets = [address for address in targets if _is_formula_cell(graph, address)]
    if not formula_targets:
        return dict(expected)
    export_graph = _graph_with_sheet_order(graph)
    pkg_name = f"parity_{uuid.uuid4().hex[:12]}"
    with TemporaryDirectory(prefix="parity_export_") as tmp:
        tmp_path = Path(tmp)
        workbook = tmp_path / "parity.xlsx"
        write_workbook(export_graph, workbook)
        document, axis_plan = _bindings_for_graph(export_graph, formula_targets)
        _write_axis_keys(workbook, axis_plan)
        bindings = validate_bindings_document(document)
        modules = generate_inverted_tree_modules(
            export_graph,
            series_bindings=bindings,
            bindings_workbook=workbook,
            blank_ranges=blank_ranges,
        )
        try:
            pkg = load_package(modules, tmp_path, name=pkg_name)
            exported: dict[str, object] = {}
            original_by_canon = {_canonical(address): address for address in formula_targets}
            for series in document["series"]:
                if "output" not in series:
                    continue
                series_id = str(series["id"])
                address = original_by_canon[_canonical(_scalar_address(series["data_range"]))]
                function = getattr(pkg, f"compute_{series_id}")
                try:
                    exported[address] = invoke_public_compute(pkg, function, {})
                except Exception as exc:
                    if not _is_export_xl_error(exc):
                        raise
                    exported[address] = _export_error_to_sentinel(exc)
            missing = [address for address in formula_targets if address not in exported]
            if missing:
                raise AssertionError(f"export package omitted formula targets {missing}")
            return {address: exported.get(address, expected[address]) for address in targets}
        finally:
            unload_package(pkg_name)


def _canonical(address: str) -> str:
    return str(canonical_address(address))


def _is_formula_cell(graph: DependencyGraph, address: str) -> bool:
    node = graph.get_node(address)
    return node is not None and node.has_formula


def _sheet_names(graph: DependencyGraph) -> list[str]:
    names: list[str] = []
    seen: set[str] = set()
    for key in graph:
        node = graph.get_node(key)
        if node is None or not node.sheet or node.sheet in seen:
            continue
        names.append(node.sheet)
        seen.add(node.sheet)
    return names


def _graph_with_sheet_order(graph: DependencyGraph) -> DependencyGraph:
    if graph.sheet_order:
        return graph
    names = _sheet_names(graph)
    if not names:
        return graph
    clone = graph.copy()
    clone.sheet_order = names
    return clone


def _bindings_for_graph(
    graph: DependencyGraph,
    formula_targets: Sequence[str],
) -> tuple[dict[str, Any], dict[str, tuple[int, str, dict[tuple[str, int], object]]]]:
    """Bind the demand cone of `formula_targets` plus covering constants.

    Leaf blocks that form a dense rectangle share one constant series so range
    consumers (INDEX, VLOOKUP, SUMPRODUCT, OFFSET) see a single covering
    catalog owner. Axis keys are written outside the used rectangle. Formula
    cells outside the cone are omitted so mixed graphs can export the emittable
    subset.
    """
    cone = _demand_cone(graph, formula_targets)
    target_set = {_canonical(address) for address in formula_targets}
    formula_nodes: list[Node] = []
    leaf_nodes: list[Node] = []
    for address in cone:
        node = graph.get_node(address)
        if node is None or not node.sheet:
            continue
        if node.has_formula:
            formula_nodes.append(node)
        else:
            leaf_nodes.append(node)

    axis_plan: dict[str, tuple[int, str, dict[tuple[str, int], object]]] = {}
    series: list[dict[str, Any]] = []
    for index, block in enumerate(_leaf_blocks(graph, leaf_nodes)):
        sheet = block[0]
        header_row, label_column = _axis_slots(graph, sheet)
        keys = axis_plan.setdefault(sheet, (header_row, label_column, {}))[2]
        series.append(
            _leaf_series_entry(
                f"leaf_{index}",
                graph,
                block,
                header_row=header_row,
                label_column=label_column,
                axis_keys=keys,
            )
        )

    for index, node in enumerate(formula_nodes):
        address = _canonical(str(node.key))
        direction = "output" if address in target_set else "internal"
        series.append(
            series_entry(
                f"cell_{index}",
                address,
                layout="scalar",
                direction=direction,
                dtype=_dtype_for_formula(node),
            )
        )
        series[-1]["sheet"] = node.sheet
    return bindings_document(*series), axis_plan


def _demand_cone(graph: DependencyGraph, formula_targets: Sequence[str]) -> set[str]:
    cone: set[str] = set()
    stack = [_canonical(address) for address in formula_targets]
    while stack:
        address = stack.pop()
        if address in cone:
            continue
        cone.add(address)
        node = graph.get_node(address)
        if node is None or not node.has_formula:
            continue
        ast = getattr(node, "formula_ast", None)
        if ast is None:
            continue
        for ref in _ast_demand_addresses(ast, address, graph):
            if ref not in cone:
                stack.append(ref)
    return cone


def _ast_demand_addresses(ast: AstNode, host: str, graph: DependencyGraph) -> list[str]:
    found: list[str] = []

    def walk(node: AstNode) -> None:
        if isinstance(node, CellRefNode):
            found.append(str(canonical_address(resolve_cell_ref(node, host))))
            return
        if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
            try:
                found.extend(str(addr) for addr in iter_ref_addresses(node, host, graph))
            except InvertedTreeExportError:
                return
            return
        if isinstance(node, FunctionCallNode):
            if normalize_excel_function_name(node.name) == "OFFSET":
                found.extend(_offset_cover_addresses(node, host, graph))
            for arg in node.args:
                walk(arg)
            return
        if isinstance(node, BinaryOpNode):
            walk(node.left)
            walk(node.right)
            return
        if isinstance(node, UnaryOpNode):
            walk(node.operand)

    walk(ast)
    return found


def _offset_cover_addresses(node: FunctionCallNode, host: str, graph: DependencyGraph) -> list[str]:
    dest = offset_index_destination(node, host)
    if dest is not None:
        try:
            return [str(addr) for addr in iter_range_addresses(dest[0], dest[1])]
        except InvertedTreeExportError:
            return []
    if len(node.args) < 3:
        return []
    rows = ast_literal_int(node.args[1])
    cols = ast_literal_int(node.args[2])
    corners = ref_window_corners(node.args[0], host)
    if corners is not None and rows is not None and cols is not None:
        shifted = shift_range_corners(corners[0], corners[1], rows, cols)
        if shifted is None:
            return []
        try:
            return [str(addr) for addr in iter_range_addresses(shifted[0], shifted[1])]
        except InvertedTreeExportError:
            return []
    sheets: set[str] = set()
    if corners is not None:
        sheets.add(parse_cell_coords(corners[0])[0])
    if not sheets:
        return []
    covered: list[str] = []
    for key in graph:
        leaf = graph.get_node(key)
        if leaf is None or leaf.has_formula or leaf.sheet not in sheets:
            continue
        covered.append(_canonical(str(leaf.key)))
    return covered


def _leaf_blocks(graph: DependencyGraph, leaves: Sequence[Node]) -> list[_LeafBlock]:
    by_sheet: dict[str, list[tuple[int, int]]] = defaultdict(list)
    for node in leaves:
        if node.sheet is None or node.row is None or not node.column:
            continue
        _sheet, row, col = parse_cell_coords(str(node.key))
        by_sheet[node.sheet].append((row, col))
    blocks: list[_LeafBlock] = []
    for sheet, coords in by_sheet.items():
        rows = [row for row, _col in coords]
        cols = [col for _row, col in coords]
        bbox = (sheet, min(rows), min(cols), max(rows), max(cols))
        if _leaf_rect_is_dense(graph, *bbox):
            blocks.append(bbox)
            continue
        by_col: dict[int, list[int]] = defaultdict(list)
        for row, col in coords:
            by_col[col].append(row)
        for col, col_rows in sorted(by_col.items()):
            for start, stop in _contiguous_runs(sorted(set(col_rows))):
                blocks.append((sheet, start, col, stop, col))
    return blocks


def _leaf_rect_is_dense(
    graph: DependencyGraph, sheet: str, row1: int, col1: int, row2: int, col2: int
) -> bool:
    for row in range(row1, row2 + 1):
        for col in range(col1, col2 + 1):
            node = graph.get_node(format_cell_key(sheet, get_column_letter(col), row))
            if node is None or node.has_formula:
                return False
    return True


def _contiguous_runs(values: Sequence[int]) -> list[tuple[int, int]]:
    if not values:
        return []
    runs: list[tuple[int, int]] = []
    start = prev = values[0]
    for value in values[1:]:
        if value == prev + 1:
            prev = value
            continue
        runs.append((start, prev))
        start = prev = value
    runs.append((start, prev))
    return runs


def _axis_slots(graph: DependencyGraph, sheet: str) -> tuple[int, str]:
    max_row = 1
    max_col = 1
    for key in graph:
        node = graph.get_node(key)
        if node is None or node.sheet != sheet:
            continue
        _sheet, row, col = parse_cell_coords(str(node.key))
        max_row = max(max_row, row)
        max_col = max(max_col, col)
    return max_row + 1, get_column_letter(max_col + 1)


def _leaf_series_entry(
    series_id: str,
    graph: DependencyGraph,
    block: _LeafBlock,
    *,
    header_row: int,
    label_column: str,
    axis_keys: dict[tuple[str, int], object],
) -> dict[str, Any]:
    sheet, row1, col1, row2, col2 = block
    start = f"{get_column_letter(col1)}{row1}"
    end = f"{get_column_letter(col2)}{row2}"
    data_range = (
        format_range_key(sheet, start, end)
        if (row1, col1) != (row2, col2)
        else (format_cell_key(sheet, get_column_letter(col1), row1))
    )
    values = []
    for row in range(row1, row2 + 1):
        for col in range(col1, col2 + 1):
            node = graph.get_node(format_cell_key(sheet, get_column_letter(col), row))
            if node is not None:
                values.append(node.value)
    dtype = _dtype_for_values(values)
    width = col2 - col1 + 1
    height = row2 - row1 + 1
    if width == 1 and height == 1:
        entry = series_entry(
            series_id, data_range, layout="scalar", direction="constant", dtype=dtype
        )
        entry["sheet"] = sheet
        return entry
    for row in range(row1, row2 + 1):
        axis_keys[("row", row)] = row
    for col in range(col1, col2 + 1):
        axis_keys[("col", col)] = col
    if width == 1:
        entry = series_entry(
            series_id,
            data_range,
            layout="series",
            direction="constant",
            dtype=dtype,
            label_column=label_column,
            key_read="int",
        )
        entry["sheet"] = sheet
        return entry
    if height == 1:
        entry = series_entry(
            series_id,
            data_range,
            layout="series",
            direction="constant",
            dtype=dtype,
            header_row=header_row,
            key_read="int",
        )
        entry["sheet"] = sheet
        return entry
    entry = series_entry(
        series_id,
        data_range,
        layout="matrix",
        direction="constant",
        dtype=dtype,
        header_row=header_row,
        label_column=label_column,
        key=("COUNTRY", "TIME_PERIOD"),
    )
    entry["sheet"] = sheet
    entry["structure"]["dimensions"] = [
        {
            "concept": "COUNTRY",
            "role": "key",
            "scope": "cell",
            "bind": {"kind": "row_label", "label_column": label_column, "read": "int"},
        },
        {
            "concept": "TIME_PERIOD",
            "role": "key",
            "scope": "cell",
            "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
        },
    ]
    return entry


def _write_axis_keys(
    workbook: Path, axis_plan: dict[str, tuple[int, str, dict[tuple[str, int], object]]]
) -> None:
    if not any(keys for _header, _label, keys in axis_plan.values()):
        return
    book = load_workbook(workbook)
    try:
        for sheet, (header_row, label_column, keys) in axis_plan.items():
            worksheet = book[sheet]
            for kind, index in keys:
                if kind == "row":
                    worksheet[f"{label_column}{index}"] = index
                else:
                    worksheet[f"{get_column_letter(index)}{header_row}"] = index
        book.save(workbook)
    finally:
        book.close()


def _dtype_for_values(values: Sequence[object]) -> str:
    kinds = {_dtype_for(value) for value in values if value is not None}
    if kinds == {"bool"}:
        return "bool"
    if kinds == {"string"}:
        return "string"
    return "float"


def _dtype_for(value: object) -> str:
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, str):
        return "string"
    if isinstance(value, datetime | date):
        return "datetime"
    if isinstance(value, XlError):
        return "float"
    return "float"


def _dtype_for_formula(node: Node) -> str:
    ast = getattr(node, "formula_ast", None)
    if ast is None:
        return "float"
    return _dtype_for_ast(ast)


def _dtype_for_ast(ast: AstNode) -> str:
    if isinstance(ast, BoolNode):
        return "bool"
    if isinstance(ast, StringNode):
        return "string"
    if isinstance(ast, BinaryOpNode):
        if ast.op in _COMPARE_OPS:
            return "bool"
        if ast.op == "&":
            return "string"
        return "float"
    if isinstance(ast, FunctionCallNode):
        name = normalize_excel_function_name(ast.name)
        if name in _BOOL_FUNCS:
            return "bool"
        if name in _STRING_FUNCS:
            return "string"
        if name == "IF" and len(ast.args) < 3:
            return "bool"
        if name in {"IFNA", "IFERROR"} and ast.args:
            kinds = {_dtype_for_ast(arg) for arg in ast.args}
            if len(kinds) == 1:
                return next(iter(kinds))
    if isinstance(ast, UnaryOpNode):
        return _dtype_for_ast(ast.operand)
    return "float"


def _scalar_address(data_range: str) -> str:
    if ":" in data_range:
        raise AssertionError(f"expected a scalar output range, got {data_range!r}")
    return data_range


def _is_export_xl_error(exc: BaseException) -> bool:
    if type(exc).__name__ not in _EXPORT_ERROR_TYPES:
        return False
    code = getattr(exc, "code", None)
    if isinstance(code, XlError):
        return True
    return isinstance(code, str) and XlError.from_text(code) is not None


def _export_error_to_sentinel(exc: BaseException) -> object:
    code = getattr(exc, "code", None)
    if isinstance(code, XlError):
        return code
    if isinstance(code, str):
        mapped = XlError.from_text(code)
        if mapped is not None:
            return mapped
    return code


def _values_match(want: object, got: object) -> bool:
    if isinstance(want, bool) or isinstance(got, bool):
        return want is got
    if isinstance(want, (int, float)) and isinstance(got, (int, float)):
        return got == pytest.approx(want)
    return want == got
