from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from math import isfinite
from typing import Any, cast

from excel_grapher import CycleError, DependencyGraph, FormulaEvaluator
from excel_grapher.evaluator.codegen import CodeGenerator


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
        for dep in graph.dependencies(addr):
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

    with FormulaEvaluator(
        graph, on_cell_evaluated=_record, blank_ranges=blank_ranges
    ) as ev:
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


def assert_code_does_not_embed_symbols(code: str, *, absent: set[str]) -> None:
    """Pruning helper: assert certain top-level runtime defs are not embedded."""
    hits = {sym for sym in absent if f"def {sym}(" in code or f"class {sym}:" in code}
    if hits:
        raise AssertionError(
            f"Expected symbols to be pruned, but found in generated code: {sorted(hits)}"
        )
