"""Unit tests for parameterized helper memoization (`xl_helper`)."""

from __future__ import annotations

import pytest

from excel_grapher.core import XlError, XlErrorException
from excel_grapher.runtime.cache import (
    CircularReferenceWarning,
    EvalContext,
    coerce_inputs_dict,
    xl_helper,
    xl_iterative_compute,
)


def _ctx() -> EvalContext:
    return EvalContext(
        inputs=coerce_inputs_dict({}),
        resolver=lambda _address: None,
    )


class TestXlHelper:
    def test_direct_xl_helper_caches_by_fn_and_kwargs(self) -> None:
        calls = {"n": 0}

        def bump(ctx: EvalContext, *, n: int) -> int:
            calls["n"] += 1
            return n * 2

        ctx = _ctx()
        assert xl_helper(ctx, bump, n=3) == 6
        assert xl_helper(ctx, bump, n=3) == 6
        assert calls["n"] == 1
        assert xl_helper(ctx, bump, n=4) == 8
        assert calls["n"] == 2

    def test_unhashable_kwargs_fail_loud(self) -> None:
        def ignore(ctx: EvalContext, *, items: list[int]) -> int:
            return len(items)

        ctx = _ctx()
        with pytest.raises(TypeError, match="hashable"):
            xl_helper(ctx, ignore, items=[1, 2])

    def test_cached_xl_error_is_re_raised(self) -> None:
        calls = {"n": 0}

        def boom(ctx: EvalContext, *, code: str) -> int:
            calls["n"] += 1
            raise XlErrorException(XlError.DIV)

        ctx = _ctx()
        with pytest.raises(XlErrorException) as first:
            xl_helper(ctx, boom, code="x")
        assert first.value.code == XlError.DIV
        assert calls["n"] == 1

        with pytest.raises(XlErrorException) as second:
            xl_helper(ctx, boom, code="x")
        assert second.value.code == XlError.DIV
        assert calls["n"] == 1

    def test_reentrant_identical_key_returns_circular_zero(self) -> None:
        def loop(ctx: EvalContext, *, n: int) -> int:
            return xl_helper(ctx, loop, n=n)

        ctx = _ctx()
        with pytest.warns(CircularReferenceWarning):
            assert xl_helper(ctx, loop, n=1) == 0


class TestHelperCacheInvalidation:
    def test_invalidate_clears_helper_cache(self) -> None:
        calls = {"n": 0}

        def once(ctx: EvalContext, *, n: int) -> int:
            calls["n"] += 1
            return n

        ctx = _ctx()
        assert xl_helper(ctx, once, n=1) == 1
        ctx.invalidate(["Inputs!A1"])
        assert xl_helper(ctx, once, n=1) == 1
        assert calls["n"] == 2

    def test_iterative_compute_restart_clears_helper_cache(self) -> None:
        calls = {"n": 0}

        def helper(ctx: EvalContext, *, n: int) -> int:
            calls["n"] += 1
            return n

        ctx = _ctx()
        ctx.iterative_enabled = True
        ctx.iterate_count = 2
        ctx.iterate_delta = 0.0  # force full iteration budget

        def target(eval_ctx: EvalContext, _address: str) -> int:
            return xl_helper(eval_ctx, helper, n=1)

        xl_iterative_compute(ctx, {"S!A1": target})
        # Each iterative restart clears helper memos, so the body runs once per pass
        # plus the final return pass.
        assert calls["n"] >= 2
