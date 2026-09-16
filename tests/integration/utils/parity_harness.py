"""Evaluator ↔ inverted-tree export helpers for function tests.

`evaluate_targets` still returns `FormulaEvaluator` results so existing
asserts keep Excel's sentinel error channel. It also emits a bindings-backed
inverted-tree package and checks those compute results against the evaluator
(#863 A1). Address-keyed `CodeGenerator.generate` was removed in #764.
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
from excel_grapher.core.types import XlError
from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.writeback import write_workbook
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    load_package,
    series_entry,
)

_LeafBlock = tuple[str, int, int, int, int]


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
    _ensure_sheet_order(graph)
    with TemporaryDirectory(prefix="parity_export_") as tmp:
        tmp_path = Path(tmp)
        workbook = tmp_path / "parity.xlsx"
        write_workbook(graph, workbook)
        document, axis_plan = _bindings_for_graph(graph, formula_targets, expected)
        _write_axis_keys(workbook, axis_plan)
        bindings = validate_bindings_document(document)
        modules = generate_inverted_tree_modules(
            graph,
            series_bindings=bindings,
            bindings_workbook=workbook,
            blank_ranges=blank_ranges,
        )
        pkg = load_package(modules, tmp_path, name=f"parity_{uuid.uuid4().hex[:12]}")
        exported: dict[str, object] = {}
        original_by_canon = {_canonical(address): address for address in formula_targets}
        for series in document["series"]:
            if "output" not in series:
                continue
            series_id = str(series["id"])
            address = original_by_canon[_canonical(_scalar_address(series["data_range"]))]
            function = getattr(pkg, f"compute_{series_id}")
            try:
                exported[address] = _as_eval_value(function())
            except Exception as exc:
                exported[address] = _as_eval_value(exc)
        missing = [address for address in formula_targets if address not in exported]
        if missing:
            raise AssertionError(f"export package omitted formula targets {missing}")
        return {address: exported.get(address, expected[address]) for address in targets}


def _canonical(address: str) -> str:
    return str(canonical_address(address))


def _is_formula_cell(graph: DependencyGraph, address: str) -> bool:
    node = graph.get_node(address)
    return node is not None and node.has_formula


def _ensure_sheet_order(graph: DependencyGraph) -> None:
    if graph.sheet_order:
        return
    names: list[str] = []
    seen: set[str] = set()
    for key in graph:
        node = graph.get_node(key)
        if node is None or not node.sheet or node.sheet in seen:
            continue
        names.append(node.sheet)
        seen.add(node.sheet)
    graph.sheet_order = names


def _bindings_for_graph(
    graph: DependencyGraph,
    formula_targets: Sequence[str],
    expected: dict[str, object],
) -> tuple[dict[str, Any], dict[str, tuple[int, str, dict[tuple[str, int], object]]]]:
    """Build scalar/series/matrix constants plus one output per formula target.

    Leaf blocks that form a dense rectangle share one constant series so range
    consumers (INDEX, VLOOKUP, SUMPRODUCT, OFFSET) see a single covering
    catalog owner. Axis keys are written outside the used rectangle.
    """
    target_set = {_canonical(address) for address in formula_targets}
    formula_nodes: list[Node] = []
    leaf_nodes: list[Node] = []
    for key in graph:
        node = graph.get_node(key)
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

    expected_by_canon = {_canonical(key): value for key, value in expected.items()}
    for index, node in enumerate(formula_nodes):
        address = _canonical(str(node.key))
        direction = "output" if address in target_set else "internal"
        dtype = _dtype_for(expected_by_canon.get(address, node.value))
        series.append(
            series_entry(
                f"cell_{index}",
                address,
                layout="scalar",
                direction=direction,
                dtype=dtype,
            )
        )
        series[-1]["sheet"] = node.sheet
    return bindings_document(*series), axis_plan


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


def _scalar_address(data_range: str) -> str:
    if ":" in data_range:
        raise AssertionError(f"expected a scalar output range, got {data_range!r}")
    return data_range


def _as_eval_value(value: object) -> object:
    code = getattr(value, "code", None)
    if isinstance(value, BaseException) and code is not None:
        if isinstance(code, XlError):
            return code
        mapped = XlError.from_text(str(code))
        return mapped if mapped is not None else code
    if isinstance(value, str):
        mapped = XlError.from_text(value)
        if mapped is not None:
            return mapped
    return value


def _values_match(want: object, got: object) -> bool:
    if isinstance(want, bool) or isinstance(got, bool):
        return want is got
    if isinstance(want, (int, float)) and isinstance(got, (int, float)):
        return got == pytest.approx(want)
    return want == got
