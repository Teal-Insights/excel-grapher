"""Evaluator helpers and export-runtime scaffold checks.

Address-keyed `CodeGenerator.generate` was removed (#764). Evaluator ↔ export
parity for packages uses inverted-tree `generate_modules` with series bindings.
`FormulaEvaluator` remains the in-process Excel engine for function tests.
"""

from __future__ import annotations

from math import isfinite
from typing import Any, cast

from excel_grapher import DependencyGraph, FormulaEvaluator


def _is_finite_number(x: object) -> bool:
    if isinstance(x, bool):
        return False
    if isinstance(x, (int, float)):
        return isfinite(float(x))
    return False


def _values_equal(a: object, b: object, *, rtol: float, atol: float) -> bool:
    if a == b:
        return True
    if _is_finite_number(a) and _is_finite_number(b):
        af = float(cast(int | float, a))
        bf = float(cast(int | float, b))
        return abs(af - bf) <= max(atol, rtol * max(abs(af), abs(bf)))
    return False


def evaluate_targets(
    graph: DependencyGraph,
    targets: list[str],
    *,
    blank_ranges: tuple[str, ...] | list[str] | None = None,
) -> dict[str, object]:
    """Evaluate `targets` with `FormulaEvaluator` and return the result map."""
    with FormulaEvaluator(graph, blank_ranges=blank_ranges) as ev:
        return cast(dict[str, object], ev.evaluate(targets))


_CACHE_EVAL_SCAFFOLD_DEFS = ("def _evaluate_address(", "def xl_cell(", "def xl_eval(")
CACHE_EVAL_SCAFFOLD_LINE_BUDGET = 80


def count_cache_eval_scaffold_lines(code: str) -> int:
    """Count lines for `_evaluate_address`, `xl_cell`, and `xl_eval` in export code."""
    lines = code.splitlines()
    try:
        start = next(
            i for i, line in enumerate(lines) if line.startswith(_CACHE_EVAL_SCAFFOLD_DEFS[0])
        )
    except StopIteration as exc:
        raise ValueError("cache eval scaffold not found in generated code") from exc

    end = len(lines)
    for index, line in enumerate(lines[start + 1 :], start + 1):
        is_next_def = line.startswith("def ") and not any(
            line.startswith(marker) for marker in _CACHE_EVAL_SCAFFOLD_DEFS
        )
        if is_next_def or line.startswith("# ---"):
            end = index
            break
    return end - start


def assert_cache_eval_scaffold_within_budget(
    code: str,
    *,
    max_lines: int = CACHE_EVAL_SCAFFOLD_LINE_BUDGET,
) -> int:
    """Assert exported cache eval helpers stay within the deduplication line budget."""
    for marker in _CACHE_EVAL_SCAFFOLD_DEFS:
        if marker not in code:
            raise AssertionError(f"Expected {marker!r} in generated export code")

    line_count = count_cache_eval_scaffold_lines(code)
    if line_count > max_lines:
        raise AssertionError(
            f"Cache eval scaffold bloated to {line_count} lines (budget {max_lines}); "
            "xl_cell/xl_eval may have diverged from _evaluate_address"
        )
    return line_count


def assert_code_does_not_embed_symbols(code: str, *, absent: set[str]) -> None:
    """Pruning helper: assert certain top-level runtime defs are not embedded."""
    hits = {sym for sym in absent if f"def {sym}(" in code or f"class {sym}:" in code}
    if hits:
        raise AssertionError(
            f"Expected symbols to be pruned, but found in generated code: {sorted(hits)}"
        )


EMBEDDED_RUNTIME_HEADER = '"""Standalone runtime for generated Excel formula code."""'
FORMULA_CELLS_MARKER = "# --- Formula cell functions ---"

DEP_TRACKING_METHOD_SYMBOLS = frozenset(
    {
        "_record_dependency",
        "invalidate",
        "set_inputs",
    }
)
DEP_TRACKING_FIELD_MARKERS = frozenset(
    {
        "deps: dict[str, set[str]]",
        "reverse_deps: dict[str, set[str]]",
    }
)
DEP_TRACKING_CALL_MARKERS = frozenset(
    {
        "ctx._record_dependency(",
    }
)

DEP_TRACKING_BASELINE_VERSION = 16
SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET = 62


def extract_embedded_runtime(code: str) -> str:
    """Return the embedded `emit_runtime` block from generated export code."""
    lines = code.splitlines()
    try:
        start = next(i for i, line in enumerate(lines) if line.strip() == EMBEDDED_RUNTIME_HEADER)
    except StopIteration as exc:
        raise ValueError("embedded runtime header not found in generated code") from exc

    end = len(lines)
    for index, line in enumerate(lines[start + 1 :], start + 1):
        if line.startswith(FORMULA_CELLS_MARKER):
            end = index
            break
        if line.startswith("DEFAULT_INPUTS = {"):
            end = index
            break
    return "\n".join(lines[start:end]).rstrip()


def count_embedded_runtime_lines(code: str) -> int:
    """Count lines in the embedded runtime block."""
    return len(extract_embedded_runtime(code).splitlines())


def dep_tracking_hits(code: str) -> dict[str, Any]:
    """Report whether dep-tracking fields, methods, and call sites are present."""
    return {
        "deps_field": any(marker in code for marker in DEP_TRACKING_FIELD_MARKERS),
        "reverse_deps_field": "reverse_deps: dict[str, set[str]]" in code,
        "methods": {symbol: f"def {symbol}(" in code for symbol in DEP_TRACKING_METHOD_SYMBOLS},
        "record_dependency_call": any(marker in code for marker in DEP_TRACKING_CALL_MARKERS),
    }


def _eval_context_class_start(lines: list[str]) -> int:
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("class EvalContext(") or stripped == "class EvalContext:":
            return index
    raise ValueError("EvalContext class not found in generated code")


def count_dep_tracking_lines(code: str) -> int:
    """Count EvalContext dep-tracking fields/methods plus `_record_dependency` call sites."""
    lines = code.splitlines()
    try:
        class_start = _eval_context_class_start(lines)
    except ValueError:
        return 0

    count = 0
    in_method = False
    method_indent = 0
    for line in lines[class_start + 1 :]:
        stripped = line.strip()
        if stripped.startswith("class ") or (
            stripped.startswith("def ") and not line.startswith("    ")
        ):
            break

        if (
            stripped.startswith("def _record_dependency(")
            or stripped.startswith("def invalidate(")
            or stripped.startswith("def set_inputs(")
        ):
            in_method = True
            method_indent = len(line) - len(line.lstrip())
            count += 1
            continue

        if in_method:
            indent = len(line) - len(line.lstrip())
            if stripped and indent <= method_indent:
                in_method = False
            else:
                count += 1
                continue

        if stripped in DEP_TRACKING_FIELD_MARKERS or stripped.startswith("stack: list[str]"):
            count += 1

    count += sum(1 for line in lines if "ctx._record_dependency(" in line)
    return count


def assert_dep_tracking_present(code: str) -> None:
    """Assert generated export embeds the invalidation subsystem."""
    hits = dep_tracking_hits(code)
    methods = cast(dict[str, bool], hits["methods"])
    missing: list[str] = []
    if not hits["deps_field"]:
        missing.append("deps field")
    if not hits["reverse_deps_field"]:
        missing.append("reverse_deps field")
    for symbol, present in methods.items():
        if not present:
            missing.append(f"def {symbol}")
    if not hits["record_dependency_call"]:
        missing.append("ctx._record_dependency call site")
    if missing:
        raise AssertionError(
            f"Expected dependency-tracking scaffold in generated code; missing: {missing}"
        )


def assert_dep_tracking_absent(code: str) -> None:
    """Assert generated export omits the invalidation subsystem."""
    hits = dep_tracking_hits(code)
    methods = cast(dict[str, bool], hits["methods"])
    present: list[str] = []
    if hits["deps_field"]:
        present.append("deps field")
    if hits["reverse_deps_field"]:
        present.append("reverse_deps field")
    for symbol, found in methods.items():
        if found:
            present.append(f"def {symbol}")
    if hits["record_dependency_call"]:
        present.append("ctx._record_dependency call site")
    if present:
        raise AssertionError(
            f"Expected dependency-tracking scaffold to be omitted; found: {present}"
        )
