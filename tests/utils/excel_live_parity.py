"""Build small workbooks, recalculate through Excel, assert evaluator ↔ live Excel parity.

Per-function live Excel tests should call `assert_evaluator_matches_live_excel` from
`tests/integration/evaluator/test_<fn>_excel_parity.py` modules (slow, run-if-available).
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

import fastpyxl.utils.cell
import pytest
import xlsxwriter

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.address_keys import normalize_key
from tests.utils.excel_workbook_parity import (
    ParityMismatchKind,
)
from tests.utils.excel_workbook_parity import (
    compare_cached_to_evaluator as _compare_cached_to_evaluator,
)
from tests.utils.modify_and_recalculate import (
    ExcelRecalculationError,
    modify_and_recalculate_workbook,
)

if TYPE_CHECKING:
    from excel_grapher import DependencyGraph


class LiveExcelParityMismatchKind(Enum):
    """Classification for live-Excel vs evaluator differences."""

    NUMERIC_DRIFT = "numeric_drift"
    XL_ERROR_CODE_MISMATCH = "xl_error_code_mismatch"
    XL_ERROR_VS_NUMBER = "xl_error_vs_number"
    NUMBER_VS_XL_ERROR = "number_vs_xl_error"
    EXCEPTION = "exception"
    NOT_IMPLEMENTED = "not_implemented"
    TYPE_MISMATCH = "type_mismatch"
    MISSING_TARGET = "missing_target"


def _workbook_kind_to_live(kind: ParityMismatchKind) -> LiveExcelParityMismatchKind:
    return LiveExcelParityMismatchKind(kind.value)


@dataclass(frozen=True, slots=True)
class LiveExcelCell:
    """One worksheet cell to populate before live recalculation."""

    address: str
    formula: str | None = None
    value: str | float | int | bool | None = None


@dataclass(frozen=True, slots=True)
class LiveExcelParityMismatch:
    address: str
    kind: LiveExcelParityMismatchKind
    excel_cached: object
    evaluator_result: object | None
    formula: str | None = None
    exception: BaseException | None = None

    def format_line(self) -> str:
        parts = [
            self.address,
            f"kind={self.kind.value}",
            f"excel_cached={self.excel_cached!r}",
            f"evaluator={self.evaluator_result!r}",
        ]
        if self.formula:
            parts.append(f"formula={self.formula[:120]}{'...' if len(self.formula) > 120 else ''}")
        if self.exception is not None:
            parts.append(f"exception={type(self.exception).__name__}: {self.exception}")
        return " | ".join(parts)


def _a1_to_row_col(a1: str) -> tuple[int, int]:
    col, row = fastpyxl.utils.cell.coordinate_from_string(a1.replace("$", ""))
    col_idx = fastpyxl.utils.cell.column_index_from_string(col) - 1
    return int(row) - 1, col_idx


def write_live_excel_workbook(
    path: Path,
    *,
    sheet: str,
    cells: tuple[LiveExcelCell, ...],
) -> None:
    """Write a one-sheet workbook with literals and formulas for live recalculation."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet(sheet)
    for cell in cells:
        row, col = _a1_to_row_col(cell.address)
        if cell.formula is not None:
            ws.write_formula(row, col, cell.formula)
        elif cell.value is not None:
            if isinstance(cell.value, str):
                ws.write_string(row, col, cell.value)
            elif isinstance(cell.value, bool):
                ws.write_boolean(row, col, cell.value)
            else:
                ws.write_number(row, col, float(cell.value))
    wb.close()


def compare_cached_to_evaluator(
    excel_cached: object,
    evaluator_result: object,
    *,
    rtol: float = 1e-5,
    atol: float = 1e-9,
) -> LiveExcelParityMismatchKind | None:
    """Return a mismatch kind when Excel cached value differs from evaluator output."""
    kind = _compare_cached_to_evaluator(
        excel_cached,
        evaluator_result,
        rtol=rtol,
        atol=atol,
    )
    if kind is None:
        return None
    return _workbook_kind_to_live(kind)


def format_live_excel_parity_report(mismatches: list[LiveExcelParityMismatch]) -> str:
    if not mismatches:
        return ""
    lines = ["Live Excel vs evaluator mismatches:"]
    for mismatch in mismatches:
        lines.append(f"  - {mismatch.format_line()}")
    return "\n".join(lines)


def compare_evaluator_to_live_excel(
    graph: DependencyGraph,
    addresses: list[str],
    *,
    rtol: float = 1e-5,
    atol: float = 1e-9,
    fail_fast: bool = False,
) -> list[LiveExcelParityMismatch]:
    """Compare evaluator output to cached values on a post-recalc dependency graph."""
    mismatches: list[LiveExcelParityMismatch] = []

    def _formula_for(addr: str) -> str | None:
        node = graph.get_node(addr)
        if node is None or node.formula is None:
            return None
        nf = node.normalized_formula
        if isinstance(nf, str) and nf.strip():
            return nf.strip()
        return node.formula

    with FormulaEvaluator(graph) as ev:
        for addr in addresses:
            node = graph.get_node(addr)
            if node is None:
                mismatches.append(
                    LiveExcelParityMismatch(
                        address=addr,
                        kind=LiveExcelParityMismatchKind.MISSING_TARGET,
                        excel_cached=None,
                        evaluator_result=None,
                    )
                )
                if fail_fast:
                    return mismatches
                continue

            formula = _formula_for(addr)
            try:
                computed = ev._evaluate_cell(addr)  # noqa: SLF001
            except NotImplementedError as exc:
                mismatches.append(
                    LiveExcelParityMismatch(
                        address=addr,
                        kind=LiveExcelParityMismatchKind.NOT_IMPLEMENTED,
                        excel_cached=node.value,
                        evaluator_result=None,
                        formula=formula,
                        exception=exc,
                    )
                )
                if fail_fast:
                    return mismatches
                continue
            except Exception as exc:
                mismatches.append(
                    LiveExcelParityMismatch(
                        address=addr,
                        kind=LiveExcelParityMismatchKind.EXCEPTION,
                        excel_cached=node.value,
                        evaluator_result=None,
                        formula=formula,
                        exception=exc,
                    )
                )
                if fail_fast:
                    return mismatches
                continue

            if computed is None and node.value is None:
                continue

            kind = compare_cached_to_evaluator(
                node.value,
                computed,
                rtol=rtol,
                atol=atol,
            )
            if kind is None:
                continue
            mismatches.append(
                LiveExcelParityMismatch(
                    address=addr,
                    kind=kind,
                    excel_cached=node.value,
                    evaluator_result=computed,
                    formula=formula,
                )
            )
            if fail_fast:
                return mismatches

    return mismatches


def assert_evaluator_matches_live_excel(
    *,
    tmp_path: Path,
    sheet: str,
    cells: tuple[LiveExcelCell, ...],
    targets: tuple[str, ...],
    rtol: float = 1e-5,
    atol: float = 1e-9,
    workbook_stem: str = "live_parity",
    fail_fast: bool = True,
) -> DependencyGraph:
    """Write a workbook, recalculate through Excel, and assert evaluator parity."""
    input_path = tmp_path / f"{workbook_stem}.xlsx"
    output_path = tmp_path / f"{workbook_stem}_recalc.xlsx"
    write_live_excel_workbook(input_path, sheet=sheet, cells=cells)

    try:
        modify_and_recalculate_workbook(input_path, output_path, {})
    except (ExcelRecalculationError, RuntimeError, ImportError) as exc:
        pytest.skip(f"Excel recalculation not available: {exc}")

    target_keys = [normalize_key(target) for target in targets]
    graph = create_dependency_graph(
        output_path,
        target_keys,
        load_values=True,
        use_cached_dynamic_refs=True,
    )

    mismatches = compare_evaluator_to_live_excel(
        graph,
        target_keys,
        rtol=rtol,
        atol=atol,
        fail_fast=fail_fast,
    )
    if mismatches:
        raise AssertionError(format_live_excel_parity_report(mismatches))
    return graph
