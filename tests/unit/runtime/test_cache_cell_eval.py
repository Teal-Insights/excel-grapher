"""Unit tests for ``xl_cell`` and ``xl_eval`` in :mod:`excel_grapher.runtime.cache`.

Covers shared evaluation paths (cache, inputs, dependencies, circular/iterative
guards) and divergent behavior (resolver miss, structural-blank ``None`` retention)
ahead of deduplicating the two entry points.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field

import pytest

from excel_grapher.core import CellValue
from excel_grapher.runtime.cache import (
    CircularReferenceWarning,
    EvalContext,
    coerce_inputs_dict,
    xl_cell,
    xl_eval,
)

CellFn = Callable[[EvalContext], CellValue]
ResolverFn = Callable[[str], CellFn | None]
_ADDR = "S!A1"
_CHILD = "S!B1"


def _missing_resolver(_address: str) -> CellFn | None:
    return None


def _ctx(
    *,
    inputs: dict[str, object] | None = None,
    resolver: ResolverFn | None = None,
    iterative_enabled: bool = False,
    iteration_values: dict[str, object] | None = None,
) -> EvalContext:
    return EvalContext(
        inputs=coerce_inputs_dict(inputs or {}),
        resolver=resolver or _missing_resolver,
        iterative_enabled=iterative_enabled,
        iteration_values=coerce_inputs_dict(iteration_values or {}),
    )


@dataclass
class CountingFn:
    """Callable cell implementation that tracks invocation count."""

    result: CellValue = 42
    calls: int = field(default=0, init=False)

    def __call__(self, _ctx: EvalContext) -> CellValue:
        self.calls += 1
        return self.result


class StructuralBlankCell:
    """Marks a formula cell as a structural blank (``None`` stays ``None``)."""

    __structural_blank__ = True

    def __call__(self, _ctx: EvalContext) -> None:
        return None


def _counting_fn(*, result: CellValue = 42) -> CountingFn:
    return CountingFn(result=result)


class TestSharedEvaluationPath:
    """Behavior common to ``xl_cell`` and ``xl_eval``."""

    @pytest.mark.parametrize(
        ("evaluate", "setup"),
        [
            pytest.param(
                lambda ctx, fn: xl_cell(ctx, _ADDR),
                lambda fn: _ctx(resolver=lambda address: fn if address == _ADDR else None),
                id="xl_cell",
            ),
            pytest.param(
                lambda ctx, fn: xl_eval(ctx, _ADDR, fn),
                lambda fn: _ctx(),
                id="xl_eval",
            ),
        ],
    )
    def test_cache_hit_skips_recompute(
        self,
        evaluate: Callable[[EvalContext, CountingFn], CellValue],
        setup: Callable[[CountingFn], EvalContext],
    ) -> None:
        fn = _counting_fn()
        ctx = setup(fn)

        assert evaluate(ctx, fn) == 42
        assert evaluate(ctx, fn) == 42
        assert fn.calls == 1

    @pytest.mark.parametrize(
        ("evaluate", "setup"),
        [
            pytest.param(
                lambda ctx, fn: xl_cell(ctx, _ADDR),
                lambda fn: _ctx(inputs={_ADDR: 5}, resolver=lambda _address: fn),
                id="xl_cell",
            ),
            pytest.param(
                lambda ctx, fn: xl_eval(ctx, _ADDR, fn),
                lambda fn: _ctx(inputs={_ADDR: 5}),
                id="xl_eval",
            ),
        ],
    )
    def test_inputs_override_resolver_and_are_cached(
        self,
        evaluate: Callable[[EvalContext, CountingFn], CellValue],
        setup: Callable[[CountingFn], EvalContext],
    ) -> None:
        fn = _counting_fn(result=99)
        ctx = setup(fn)

        assert evaluate(ctx, fn) == 5
        assert fn.calls == 0
        assert ctx.cache[_ADDR] == 5

    def test_xl_cell_records_dependency_from_active_stack(self) -> None:
        def parent_fn(ctx: EvalContext) -> CellValue:
            xl_cell(ctx, _CHILD)
            return 1

        ctx = _ctx(
            inputs={_CHILD: 10},
            resolver=lambda address: parent_fn if address == _ADDR else None,
        )
        xl_cell(ctx, _ADDR)

        assert _CHILD in ctx.deps[_ADDR]

    def test_xl_eval_records_dependency_from_active_stack(self) -> None:
        def parent_fn(ctx: EvalContext) -> CellValue:
            xl_eval(ctx, _CHILD, lambda _c: 0)
            return 1

        ctx = _ctx(resolver=lambda address: parent_fn if address == _ADDR else None)
        xl_cell(ctx, _ADDR)

        assert _CHILD in ctx.deps[_ADDR]

    @pytest.mark.parametrize(
        ("evaluate", "setup"),
        [
            pytest.param(
                lambda ctx, fn: xl_cell(ctx, _ADDR),
                lambda fn: _ctx(resolver=lambda address: fn if address == _ADDR else None),
                id="xl_cell",
            ),
            pytest.param(
                lambda ctx, fn: xl_eval(ctx, _ADDR, fn),
                lambda fn: _ctx(),
                id="xl_eval",
            ),
        ],
    )
    def test_non_iterative_cycle_returns_zero_with_warning(
        self,
        evaluate: Callable[[EvalContext, CellFn], CellValue],
        setup: Callable[[CellFn], EvalContext],
    ) -> None:
        def self_ref(ctx: EvalContext) -> CellValue:
            return evaluate(ctx, self_ref)

        ctx = setup(self_ref)

        with pytest.warns(CircularReferenceWarning):
            assert evaluate(ctx, self_ref) == 0

    @pytest.mark.parametrize("entrypoint", ["xl_cell", "xl_eval"])
    def test_iterative_mode_uses_iteration_values_while_computing(
        self,
        entrypoint: str,
    ) -> None:
        ctx = _ctx(
            iterative_enabled=True,
            iteration_values={_ADDR: 99},
        )
        ctx.computing.add(_ADDR)

        if entrypoint == "xl_cell":
            assert xl_cell(ctx, _ADDR) == 99
        else:
            assert xl_eval(ctx, _ADDR, lambda _c: 0) == 99


class TestXlCellDivergentBehavior:
    """Behavior specific to ``xl_cell`` (resolver lookup and structural blanks)."""

    def test_resolver_miss_raises_key_error(self) -> None:
        ctx = _ctx(resolver=_missing_resolver)

        with pytest.raises(KeyError, match=r"Cell S!Z99 not found in graph"):
            xl_cell(ctx, "S!Z99")

    def test_structural_blank_retains_none(self) -> None:
        blank_fn = StructuralBlankCell()
        ctx = _ctx(resolver=lambda address: blank_fn if address == "S!A2" else None)

        assert xl_cell(ctx, "S!A2") is None
        assert ctx.cache["S!A2"] is None

    def test_non_structural_none_normalized_to_zero(self) -> None:
        fn = _counting_fn(result=None)
        ctx = _ctx(resolver=lambda address: fn if address == _ADDR else None)

        assert xl_cell(ctx, _ADDR) == 0
        assert ctx.cache[_ADDR] == 0


class TestXlEvalDivergentBehavior:
    """Behavior specific to ``xl_eval`` (direct fn, no resolver)."""

    def test_does_not_call_resolver(self) -> None:
        def bad_resolver(_address: str) -> CellFn:
            raise AssertionError("resolver must not be called")

        def fn(_ctx: EvalContext) -> int:
            return 7

        ctx = _ctx(resolver=bad_resolver)
        assert xl_eval(ctx, _ADDR, fn) == 7

    def test_none_result_always_normalized_to_zero(self) -> None:
        def returns_none(_ctx: EvalContext) -> None:
            return None

        ctx = _ctx()
        assert xl_eval(ctx, _ADDR, returns_none) == 0
        assert ctx.cache[_ADDR] == 0

    def test_structural_blank_flag_still_normalizes_none_to_zero(self) -> None:
        blank_fn = StructuralBlankCell()
        ctx = _ctx()
        assert xl_eval(ctx, "S!A2", blank_fn) == 0
        assert ctx.cache["S!A2"] == 0
