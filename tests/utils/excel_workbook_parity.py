"""Compare FormulaEvaluator output to **Excel workbook cached values** on a DependencyGraph.

This is one **evaluator ↔ Excel** path: last-saved results in the file (no live
recalc). For parity against **live Excel**, use xlwings / COM automation when
available (see `modify_and_recalculate_workbook` and `test_golden_master`).

Used by slow LIC-DSF tests for triage reporting and assertions.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import TYPE_CHECKING

from excel_grapher import FormulaEvaluator, XlError

if TYPE_CHECKING:
    from excel_grapher import DependencyGraph


class ParityMismatchKind(Enum):
    """High-level classification for workbook-vs-evaluator differences."""

    NUMERIC_DRIFT = "numeric_drift"
    XL_ERROR_CODE_MISMATCH = "xl_error_code_mismatch"
    XL_ERROR_VS_NUMBER = "xl_error_vs_number"
    NUMBER_VS_XL_ERROR = "number_vs_xl_error"
    EXCEPTION = "exception"
    NOT_IMPLEMENTED = "not_implemented"
    NONE_RESULT = "none_result"
    TYPE_MISMATCH = "type_mismatch"


@dataclass(frozen=True, slots=True)
class WorkbookParityMismatch:
    address: str
    kind: ParityMismatchKind
    excel_cached: object
    evaluator_result: object | None
    formula: str | None
    exception: BaseException | None = None

    def format_line(self) -> str:
        parts = [
            f"{self.address}",
            f"kind={self.kind.value}",
            f"excel_cached={self.excel_cached!r}",
            f"evaluator={self.evaluator_result!r}",
        ]
        if self.formula:
            parts.append(f"formula={self.formula[:120]}{'...' if len(self.formula) > 120 else ''}")
        if self.exception is not None:
            parts.append(f"exception={type(self.exception).__name__}: {self.exception}")
        return " | ".join(parts)


def _numeric_close(a: float, b: float, *, rtol: float, atol: float) -> bool:
    if a == b:
        return True
    scale = max(abs(a), abs(b), 1.0)
    return abs(a - b) <= max(atol, rtol * scale)


def _normalize_cached_excel_value(raw: object) -> object:
    if isinstance(raw, str):
        err = XlError.from_text(raw)
        if err is not None:
            return err
    return raw


def compare_cached_to_evaluator(
    excel_cached: object,
    evaluator_result: object,
    *,
    rtol: float = 1e-5,
    atol: float = 1e-9,
) -> ParityMismatchKind | None:
    """Return a mismatch kind when Excel cached value differs from evaluator output."""
    cached = _normalize_cached_excel_value(excel_cached)

    if isinstance(cached, XlError):
        if isinstance(evaluator_result, XlError):
            if cached == evaluator_result:
                return None
            return ParityMismatchKind.XL_ERROR_CODE_MISMATCH
        return ParityMismatchKind.XL_ERROR_VS_NUMBER

    if isinstance(evaluator_result, XlError):
        return ParityMismatchKind.NUMBER_VS_XL_ERROR

    if (
        isinstance(cached, (int, float))
        and not isinstance(cached, bool)
        and isinstance(evaluator_result, (int, float))
        and not isinstance(evaluator_result, bool)
    ):
        if _numeric_close(float(cached), float(evaluator_result), rtol=rtol, atol=atol):
            return None
        return ParityMismatchKind.NUMERIC_DRIFT

    if cached == evaluator_result:
        return None
    return ParityMismatchKind.TYPE_MISMATCH


def _is_comparable_cached_value(cached: object) -> bool:
    if isinstance(cached, XlError):
        return True
    if isinstance(cached, (int, float)) and not isinstance(cached, bool):
        return True
    return isinstance(cached, str) and XlError.from_text(cached) is not None


def compare_evaluator_to_excel_cache(
    graph: DependencyGraph,
    addresses: list[str],
    *,
    rtol: float = 1e-5,
    atol: float = 1e-9,
    fail_fast: bool = False,
) -> list[WorkbookParityMismatch]:
    """Evaluate each address and compare to the node's cached workbook value.

    Compares finite numeric cells and Excel error strings (e.g. ``'#NUM!'``).
    """
    mismatches: list[WorkbookParityMismatch] = []

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
                continue
            cached = node.value
            formula = _formula_for(addr)
            try:
                computed = ev._evaluate_cell(addr)  # noqa: SLF001
            except NotImplementedError as e:
                m = WorkbookParityMismatch(
                    address=addr,
                    kind=ParityMismatchKind.NOT_IMPLEMENTED,
                    excel_cached=cached,
                    evaluator_result=None,
                    formula=formula,
                    exception=e,
                )
                mismatches.append(m)
                if fail_fast:
                    return mismatches
                continue
            except Exception as e:
                m = WorkbookParityMismatch(
                    address=addr,
                    kind=ParityMismatchKind.EXCEPTION,
                    excel_cached=cached,
                    evaluator_result=None,
                    formula=formula,
                    exception=e,
                )
                mismatches.append(m)
                if fail_fast:
                    return mismatches
                continue

            if computed is None and cached is None:
                continue

            if not _is_comparable_cached_value(cached):
                continue

            kind = compare_cached_to_evaluator(
                cached,
                computed,
                rtol=rtol,
                atol=atol,
            )
            if kind is None:
                continue
            mismatches.append(
                WorkbookParityMismatch(
                    address=addr,
                    kind=kind,
                    excel_cached=cached,
                    evaluator_result=computed,
                    formula=formula,
                )
            )
            if fail_fast:
                return mismatches

    return mismatches


def format_workbook_parity_report(mismatches: list[WorkbookParityMismatch]) -> str:
    if not mismatches:
        return ""
    lines = ["Workbook vs evaluator mismatches:"]
    for m in mismatches:
        lines.append(f"  - {m.format_line()}")
    return "\n".join(lines)


def assert_workbook_parity(
    graph: DependencyGraph,
    addresses: list[str],
    *,
    rtol: float = 1e-5,
    atol: float = 1e-9,
    fail_fast: bool = True,
) -> None:
    mismatches = compare_evaluator_to_excel_cache(
        graph, addresses, rtol=rtol, atol=atol, fail_fast=fail_fast
    )
    if mismatches:
        raise AssertionError(format_workbook_parity_report(mismatches))
