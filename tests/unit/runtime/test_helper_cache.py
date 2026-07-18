"""Unit tests for parameterized helper memoization (``xl_helper`` / ``xl_memoize``)."""

from __future__ import annotations

from typing import cast

import pytest

from excel_grapher.core import XlError, XlErrorException
from excel_grapher.runtime.cache import (
    CircularReferenceWarning,
    EvalContext,
    coerce_inputs_dict,
    xl_helper,
    xl_iterative_compute,
    xl_memoize,
)


def _ctx() -> EvalContext:
    return EvalContext(
        inputs=coerce_inputs_dict({}),
        resolver=lambda _address: None,
    )


class TestXlMemoizeWarmContext:
    """MCVE: period-recurrence helpers share memos under one warm context."""

    def test_accum_recurrence_hits_cache_on_adjacent_years(self) -> None:
        calls = {"n": 0}

        @xl_memoize
        def accum(ctx: EvalContext, *, time_period: int) -> int:
            calls["n"] += 1
            if time_period <= 1:
                return 1
            return accum(ctx, time_period=time_period - 1) + 1

        ctx = _ctx()
        assert accum(ctx, time_period=50) == 50
        assert calls["n"] == 50

        # Warm context: next year must be one body entry, not another full walk.
        assert accum(ctx, time_period=51) == 51
        assert calls["n"] == 51

        assert accum(ctx, time_period=50) == 50
        assert calls["n"] == 51  # cache hit

        # Invalidation must drop helper memos.
        ctx.set_inputs(coerce_inputs_dict({"Inputs!A1": 1}))
        calls["n"] = 0
        assert accum(ctx, time_period=50) == 50
        assert calls["n"] == 50


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


class TestXlMemoizeBinding:
    def test_positional_args_after_ctx_are_bound(self) -> None:
        calls = {"n": 0}

        @xl_memoize
        def add(ctx: EvalContext, left: int, right: int = 0) -> int:
            calls["n"] += 1
            return left + right

        ctx = _ctx()
        assert add(ctx, 2, 3) == 5
        assert add(ctx, 2, right=3) == 5
        assert calls["n"] == 1

    def test_decorated_name_shares_cache_with_xl_helper(self) -> None:
        calls = {"n": 0}

        def body(ctx: EvalContext, *, time_period: int) -> int:
            calls["n"] += 1
            if time_period <= 1:
                return 1
            return cast(int, xl_helper(ctx, body, time_period=time_period - 1)) + 1

        memoized = xl_memoize(body)
        ctx = _ctx()
        assert memoized(ctx, time_period=10) == 10
        assert calls["n"] == 10
        # Direct xl_helper on the same underlying fn shares the memo table.
        assert xl_helper(ctx, body, time_period=10) == 10
        assert calls["n"] == 10


class TestHelperCacheInvalidation:
    def test_invalidate_clears_helper_cache(self) -> None:
        calls = {"n": 0}

        @xl_memoize
        def once(ctx: EvalContext, *, n: int) -> int:
            calls["n"] += 1
            return n

        ctx = _ctx()
        assert once(ctx, n=1) == 1
        ctx.invalidate(["Inputs!A1"])
        assert once(ctx, n=1) == 1
        assert calls["n"] == 2

    def test_iterative_compute_restart_clears_helper_cache(self) -> None:
        calls = {"n": 0}

        @xl_memoize
        def helper(ctx: EvalContext, *, n: int) -> int:
            calls["n"] += 1
            return n

        ctx = _ctx()
        ctx.iterative_enabled = True
        ctx.iterate_count = 2
        ctx.iterate_delta = 0.0  # force full iteration budget

        def target(eval_ctx: EvalContext, _address: str) -> int:
            return helper(eval_ctx, n=1)

        xl_iterative_compute(ctx, {"S!A1": target})
        # Each iterative restart clears helper memos, so the body runs once per pass
        # plus the final return pass.
        assert calls["n"] >= 2
