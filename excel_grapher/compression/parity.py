"""Round-trip value parity between original and expanded AST maps."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from math import isfinite
from typing import cast

from excel_grapher.core.address_keys import normalize_key, parse_address
from excel_grapher.core.formula_ast import AstNode
from excel_grapher.core.types import CellValue, XlError
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node

from .expand import expand_compressed_to_cells
from .types import CompressedNode

_PARITY_FORMULA_PREFIX = "=__compression_parity__!"


@dataclass(frozen=True, slots=True)
class CompressionParityMismatch:
    """One cell where original and expanded evaluation differ."""

    cell_key: str
    original_value: object
    expanded_value: object


def assert_compression_parity(
    original: Mapping[str, AstNode],
    compressed: Mapping[str, CompressedNode],
    *,
    input_values: Mapping[str, CellValue],
    rtol: float = 1e-9,
    atol: float = 0.0,
) -> None:
    """Assert expanded compressed ASTs evaluate to the same values as originals.

    Args:
        original: Per-cell AST map before compression.
        compressed: Mixed compressed map to expand and evaluate.
        input_values: Leaf cell values shared by both evaluation graphs.
        rtol: Relative tolerance for finite float comparison.
        atol: Absolute tolerance for finite float comparison.

    Raises:
        AssertionError: When cell keys differ or any target value mismatches.
    """
    mismatches = compare_compression_parity(
        original,
        compressed,
        input_values=input_values,
        rtol=rtol,
        atol=atol,
    )
    if not mismatches:
        return
    lines = [
        f"{item.cell_key}: original={item.original_value!r} expanded={item.expanded_value!r}"
        for item in mismatches
    ]
    raise AssertionError("Compression parity failed:\n" + "\n".join(lines))


def compare_compression_parity(
    original: Mapping[str, AstNode],
    compressed: Mapping[str, CompressedNode],
    *,
    input_values: Mapping[str, CellValue],
    rtol: float = 1e-9,
    atol: float = 0.0,
) -> list[CompressionParityMismatch]:
    """Compare original and expanded compressed AST evaluation cell by cell."""
    expanded = expand_compressed_to_cells(compressed)
    original_keys = {normalize_key(key) for key in original}
    expanded_keys = set(expanded)
    if original_keys != expanded_keys:
        missing = sorted(original_keys - expanded_keys)
        extra = sorted(expanded_keys - original_keys)
        raise AssertionError(
            f"Expanded cell keys do not match original (missing={missing!r}, extra={extra!r})"
        )

    normalized_original = {normalize_key(key): ast for key, ast in original.items()}
    original_results = _evaluate_ast_map(normalized_original, input_values)
    expanded_results = _evaluate_ast_map(expanded, input_values)

    mismatches: list[CompressionParityMismatch] = []
    for cell_key in sorted(original_keys):
        left = original_results[cell_key]
        right = expanded_results[cell_key]
        if not compression_values_equal(left, right, rtol=rtol, atol=atol):
            mismatches.append(
                CompressionParityMismatch(
                    cell_key=cell_key,
                    original_value=left,
                    expanded_value=right,
                )
            )
    return mismatches


def compression_values_equal(
    left: object,
    right: object,
    *,
    rtol: float = 1e-9,
    atol: float = 0.0,
) -> bool:
    """Return True when two evaluator results match within tolerance."""
    if left == right:
        return True
    if isinstance(left, XlError) and isinstance(right, XlError):
        return left == right
    if _is_finite_number(left) and _is_finite_number(right):
        lf = float(cast(int | float, left))
        rf = float(cast(int | float, right))
        return abs(lf - rf) <= max(atol, rtol * max(abs(lf), abs(rf)))
    return False


def _evaluate_ast_map(
    ast_map: Mapping[str, AstNode],
    input_values: Mapping[str, CellValue],
) -> dict[str, CellValue]:
    graph = _build_graph_from_ast_map(ast_map, input_values)
    targets = list(ast_map)
    with FormulaEvaluator(graph) as evaluator:
        results = evaluator.evaluate(targets)
    if isinstance(results, dict):
        return results
    assert len(targets) == 1
    return {targets[0]: results}


def _build_graph_from_ast_map(
    ast_map: Mapping[str, AstNode],
    input_values: Mapping[str, CellValue],
) -> DependencyGraph:
    graph = DependencyGraph()
    preparsed: dict[str, AstNode] = {}

    for address, value in input_values.items():
        graph.add_node(_leaf_node(address, value))

    for cell_key, ast in ast_map.items():
        normalized_key = normalize_key(cell_key)
        formula_key = _parity_formula_key(normalized_key)
        graph.add_node(_formula_node(normalized_key, formula_key))
        preparsed[formula_key] = ast

    graph.preparsed_formulas = preparsed
    return graph


def _parity_formula_key(cell_key: str) -> str:
    return f"{_PARITY_FORMULA_PREFIX}{cell_key}"


def _leaf_node(address: str, value: CellValue) -> Node:
    normalized_key = normalize_key(address)
    sheet, coord = parse_address(normalized_key)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=None,
        normalized_formula=None,
        value=value,
        is_leaf=True,
    )


def _formula_node(address: str, formula_key: str) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula_key,
        normalized_formula=formula_key,
        value=None,
        is_leaf=False,
    )


def _is_finite_number(value: object) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return isfinite(float(value))
    return False
