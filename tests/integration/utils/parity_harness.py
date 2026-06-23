"""Evaluator ↔ export parity: `FormulaEvaluator` vs generated standalone code.

Excel reference checks live elsewhere (e.g. `excel_workbook_parity` for cached
workbook values; live Excel via automation when available). See `.cursor/rules/parity.mdc`.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from math import isfinite
from typing import Any, cast

from excel_grapher import CycleError, DependencyGraph, FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator


@dataclass(frozen=True, slots=True)
class ParityResult:
    evaluator_results: dict[str, object]
    generated_results: dict[str, object]
    generated_code: str


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


def _dependency_closure(graph: DependencyGraph, targets: list[str]) -> set[str]:
    closure: set[str] = set()
    stack = list(targets)
    while stack:
        addr = stack.pop()
        if addr in closure:
            continue
        if graph.get_node(addr) is None:
            continue
        closure.add(addr)
        for dep in graph.get_dependencies(addr):
            if graph.get_node(dep) is None:
                continue
            stack.append(dep)
    return closure


def _dependency_order(graph: DependencyGraph, targets: list[str]) -> list[str]:
    closure = _dependency_closure(graph, targets)
    if not closure:
        return list(targets)
    try:
        eval_order = graph.evaluation_order(strict=False)
    except CycleError:
        eval_order = []
    ordered = [addr for addr in eval_order if addr in closure]
    missing = [addr for addr in closure if addr not in ordered]
    if missing:
        ordered.extend(sorted(missing))
    return ordered


def exec_generated_code(
    graph: DependencyGraph,
    targets: list[str],
    *,
    namespace_seed: dict[str, object] | None = None,
    blank_ranges: list[str] | tuple[str, ...] | None = None,
) -> tuple[dict[str, object], str, dict[str, object]]:
    """Generate + exec code for targets and return (results, code, namespace)."""
    code = CodeGenerator(graph).generate(targets, blank_ranges=blank_ranges)
    ns: dict[str, object] = dict(namespace_seed or {})
    exec(code, ns)
    compute_all = ns["compute_all"]
    assert callable(compute_all)
    compute_all_typed = cast(Callable[[], dict[str, object]], compute_all)
    generated_results = compute_all_typed()
    assert isinstance(generated_results, dict)
    return generated_results, code, ns


def exec_generated_code_with_cache(
    graph: DependencyGraph,
    targets: list[str],
    *,
    namespace_seed: dict[str, object] | None = None,
    blank_ranges: list[str] | tuple[str, ...] | None = None,
) -> tuple[dict[str, object], str, dict[str, object]]:
    """Generate + exec code for targets and return (cache, code, namespace)."""
    code = CodeGenerator(graph).generate(targets, blank_ranges=blank_ranges)
    ns: dict[str, object] = dict(namespace_seed or {})
    exec(code, ns)
    merged = dict(cast(dict[str, object], ns["DEFAULT_INPUTS"]))
    resolver = cast(Callable[[str], object], ns["_resolve_formula"])
    ctx = cast(Callable[..., object], ns["EvalContext"])(inputs=merged, resolver=resolver)
    xl_cell = cast(Callable[..., object], ns["xl_cell"])
    for target in targets:
        xl_cell(ctx, target)
    ctx_any = cast(Any, ctx)
    cache = cast(dict[str, object], ctx_any.cache)
    return dict(cache), code, ns


def assert_codegen_matches_evaluator(
    graph: DependencyGraph,
    targets: list[str],
    *,
    rtol: float = 0.0,
    atol: float = 0.0,
    dependency_order: bool = False,
    fail_fast: bool = False,
    blank_ranges: tuple[str, ...] | None = None,
) -> ParityResult:
    """Assert evaluator results match generated code for the given targets."""
    compare_targets = _dependency_order(graph, targets) if dependency_order else list(targets)
    eval_computed: dict[str, object] = {}

    def _record(address: str, value: object) -> None:
        eval_computed[address] = value

    with FormulaEvaluator(graph, on_cell_evaluated=_record, blank_ranges=blank_ranges) as ev:
        evaluator_results = cast(dict[str, object], ev.evaluate(targets))

    generated_cache, code, _ns = exec_generated_code_with_cache(
        graph, targets, blank_ranges=blank_ranges
    )
    generated_results = {t: generated_cache[t] for t in targets}

    missing = [t for t in targets if t not in evaluator_results or t not in generated_results]
    if missing:
        raise AssertionError(f"Missing targets in results: {missing}")

    mismatches: list[tuple[str, object, object]] = []
    for idx, t in enumerate(compare_targets):
        ev_val = eval_computed.get(t)
        gen_val = generated_cache.get(t)
        if ev_val is None or gen_val is None:
            continue
        if not _values_equal(ev_val, gen_val, rtol=rtol, atol=atol):
            if fail_fast:
                node = graph.get_node(t)
                formula = None if node is None else node.formula
                normalized = None if node is None else node.normalized_formula
                detail_parts: list[str] = []
                if formula:
                    detail_parts.append(f"formula={formula}")
                if normalized and normalized != formula:
                    detail_parts.append(f"normalized_formula={normalized}")
                kind = (
                    "numeric_drift"
                    if (_is_finite_number(ev_val) and _is_finite_number(gen_val))
                    else "value_mismatch"
                )
                detail = (" (" + "; ".join(detail_parts) + ")") if detail_parts else ""
                raise AssertionError(
                    f"First parity mismatch ({kind}) at "
                    f"{t}{detail} [{idx + 1}/{len(compare_targets)}]: "
                    f"evaluator={ev_val!r} generated={gen_val!r}"
                )
            mismatches.append((t, ev_val, gen_val))

    if mismatches:
        lines = ["Parity mismatch (evaluator vs generated):"]
        for t, ev_val, gen_val in mismatches[:25]:
            lines.append(f"- {t}: evaluator={ev_val!r} generated={gen_val!r}")
        if len(mismatches) > 25:
            lines.append(f"... plus {len(mismatches) - 25} more mismatches")
        raise AssertionError("\n".join(lines))

    return ParityResult(
        evaluator_results=evaluator_results,
        generated_results=generated_results,
        generated_code=code,
    )


_CACHE_EVAL_SCAFFOLD_DEFS = ("def _evaluate_address(", "def xl_cell(", "def xl_eval(")
# Pre-refactor standalone exports embedded ~78 lines for xl_cell + xl_eval alone.
# Post-refactor shared helper keeps the block at ~69 lines; budget guards re-bloat.
CACHE_EVAL_SCAFFOLD_LINE_BUDGET = 72


def count_cache_eval_scaffold_lines(code: str) -> int:
    """Count lines for ``_evaluate_address``, ``xl_cell``, and ``xl_eval`` in export code."""
    lines = code.splitlines()
    try:
        start = next(
            i for i, line in enumerate(lines) if line.startswith(_CACHE_EVAL_SCAFFOLD_DEFS[0])
        )
    except StopIteration as exc:
        raise ValueError("cache eval scaffold not found in generated code") from exc

    end = len(lines)
    for index, line in enumerate(lines[start + 1 :], start + 1):
        if line.startswith("def ") and not any(
            line.startswith(marker) for marker in _CACHE_EVAL_SCAFFOLD_DEFS
        ):
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

# Baseline for non-iterative minimal export (S!A1 leaf + S!B1 formula).
DEP_TRACKING_BASELINE_VERSION = 4
SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET = 54


def extract_embedded_runtime(code: str) -> str:
    """Return the embedded ``emit_runtime`` block from generated export code."""
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
    """Count EvalContext dep-tracking fields/methods plus ``_record_dependency`` call sites."""
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
    """Assert generated export omits the invalidation subsystem (Sprint 2 target)."""
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
