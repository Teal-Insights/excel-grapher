"""Range watches for export invalidation (#583).

Leaf-safe `xl_range` / OFFSET result rectangles record one geometric watch
instead of exploding `reverse_deps`. Mixed formula interiors stay on per-cell
recording (fail closed).
"""

from __future__ import annotations

from collections.abc import Callable

from excel_grapher.core import CellValue
from excel_grapher.exporter.export_runtime.math import xl_sum
from excel_grapher.exporter.export_runtime.offset import xl_offset, xl_range
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict, xl_cell
from excel_grapher.runtime.cache_eval_slim import EvalContext as SlimEvalContext

CellFn = Callable[[EvalContext], CellValue]
ResolverFn = Callable[[str], CellFn | None]


def _ctx(
    inputs: dict[str, object],
    formulas: dict[str, CellFn],
) -> EvalContext:
    def resolver(address: str) -> CellFn | None:
        return formulas.get(address)

    return EvalContext(inputs=coerce_inputs_dict(inputs), resolver=resolver)


def _reverse_parents(ctx: EvalContext, child: str) -> set[str]:
    return set(ctx.reverse_deps.get(child, ()))


def test_static_sum_range_records_one_watch_not_per_cell_deps() -> None:
    n = 200
    inputs = {f"S!A{row}": float(row) for row in range(1, n + 1)}

    def sum_col(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_range(ctx, f"S!A1:A{n}"))

    ctx = _ctx(inputs, {"S!B1": sum_col})
    assert xl_cell(ctx, "S!B1") == float(n * (n + 1) / 2)

    assert ctx.range_watches["S!B1"] == [("S", 1, 1, n, 1)]
    assert "S!A1" not in ctx.reverse_deps
    assert f"S!A{n}" not in ctx.reverse_deps
    assert ctx.deps.get("S!B1", set()) == set()


def test_one_by_one_range_is_still_a_rectangle_watch() -> None:
    def sum_a1(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_range(ctx, "S!A1:A1"))

    ctx = _ctx({"S!A1": 4.0}, {"S!B1": sum_a1})
    assert xl_cell(ctx, "S!B1") == 4.0
    assert ctx.range_watches["S!B1"] == [("S", 1, 1, 1, 1)]
    assert "S!A1" not in ctx.reverse_deps


def test_set_inputs_inside_watched_rect_invalidates_and_recomputes() -> None:
    n = 50
    inputs = {f"S!A{row}": 1.0 for row in range(1, n + 1)}
    calls = {"n": 0}

    def sum_col(ctx: EvalContext) -> CellValue:
        calls["n"] += 1
        return xl_sum(xl_range(ctx, f"S!A1:A{n}"))

    ctx = _ctx(inputs, {"S!B1": sum_col})
    assert xl_cell(ctx, "S!B1") == 50.0
    assert calls["n"] == 1

    ctx.set_inputs({"S!A25": 11.0})
    assert "S!B1" not in ctx.cache
    assert "S!B1" not in ctx.range_watches
    assert xl_cell(ctx, "S!B1") == 60.0
    assert calls["n"] == 2
    assert ctx.range_watches["S!B1"] == [("S", 1, 1, n, 1)]


def test_set_inputs_outside_watched_rect_leaves_cache_warm() -> None:
    def sum_a1_a2(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_range(ctx, "S!A1:A2"))

    ctx = _ctx({"S!A1": 1.0, "S!A2": 2.0, "S!A3": 3.0}, {"S!B1": sum_a1_a2})
    assert xl_cell(ctx, "S!B1") == 3.0
    ctx.set_inputs({"S!A3": 99.0})
    assert "S!B1" in ctx.cache
    assert xl_cell(ctx, "S!B1") == 3.0


def test_blank_inside_rect_is_still_a_precedent() -> None:
    """Sparse occupancy does not shrink watch geometry (#579 / #583)."""

    def sum_a1_a3(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_range(ctx, "S!A1:A3"))

    ctx = _ctx({"S!A1": 1.0, "S!A3": 3.0}, {"S!B1": sum_a1_a3})
    assert xl_cell(ctx, "S!B1") == 4.0
    assert ctx.range_watches["S!B1"] == [("S", 1, 1, 3, 1)]

    ctx.set_inputs({"S!A2": 10.0})
    assert xl_cell(ctx, "S!B1") == 14.0


def test_multi_area_sum_does_not_watch_the_hole() -> None:
    def sum_gapped(ctx: EvalContext) -> CellValue:
        return xl_sum(
            xl_cell(ctx, "S!A1"),
            xl_range(ctx, "S!A3:A5"),
            xl_range(ctx, "S!A6:A7"),
        )

    ctx = _ctx(
        {
            "S!A1": 1.0,
            "S!A2": 100.0,
            "S!A3": 3.0,
            "S!A4": 4.0,
            "S!A5": 5.0,
            "S!A6": 6.0,
            "S!A7": 7.0,
        },
        {"S!B1": sum_gapped},
    )
    assert xl_cell(ctx, "S!B1") == 26.0
    assert ctx.range_watches["S!B1"] == [("S", 3, 1, 7, 1)]
    assert _reverse_parents(ctx, "S!A1") == {"S!B1"}
    assert "S!A2" not in ctx.reverse_deps

    ctx.set_inputs({"S!A2": 999.0})
    assert "S!B1" in ctx.cache
    assert xl_cell(ctx, "S!B1") == 26.0

    ctx.set_inputs({"S!A4": 40.0})
    assert xl_cell(ctx, "S!B1") == 62.0


def test_overlapping_areas_coalesce_to_union() -> None:
    def sum_overlap(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_range(ctx, "S!A1:A3"), xl_range(ctx, "S!A2:A4"))

    ctx = _ctx(
        {"S!A1": 1.0, "S!A2": 2.0, "S!A3": 3.0, "S!A4": 4.0},
        {"S!B1": sum_overlap},
    )
    # Excel double-counts the overlap; watches merge to the union.
    assert xl_cell(ctx, "S!B1") == 15.0
    assert ctx.range_watches["S!B1"] == [("S", 1, 1, 4, 1)]


def test_mixed_formula_interior_falls_back_to_per_cell_deps() -> None:
    def a3(ctx: EvalContext) -> CellValue:
        return float(xl_cell(ctx, "S!A1")) + 1.0

    def sum_mixed(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_range(ctx, "S!A1:A3"))

    ctx = _ctx({"S!A1": 1.0, "S!A2": 2.0}, {"S!A3": a3, "S!B1": sum_mixed})
    assert xl_cell(ctx, "S!B1") == 5.0
    assert ctx.range_watches.get("S!B1", []) == []
    assert ctx.deps["S!B1"] == {"S!A1", "S!A2", "S!A3"}
    assert _reverse_parents(ctx, "S!A2") == {"S!B1"}

    ctx.set_inputs({"S!A2": 20.0})
    assert xl_cell(ctx, "S!B1") == 23.0


def test_formula_in_hole_is_not_a_false_dep() -> None:
    def hole_formula(ctx: EvalContext) -> CellValue:
        return float(xl_cell(ctx, "S!B1")) + 1.0

    def sum_gapped(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_cell(ctx, "S!A1"), xl_range(ctx, "S!A3:A4"))

    ctx = _ctx(
        {"S!A1": 1.0, "S!A3": 3.0, "S!A4": 4.0},
        {"S!A2": hole_formula, "S!B1": sum_gapped},
    )
    assert xl_cell(ctx, "S!B1") == 8.0
    assert xl_cell(ctx, "S!A2") == 9.0
    assert "S!A2" not in ctx.deps.get("S!B1", set())
    watches = ctx.range_watches.get("S!B1", [])
    for _sheet, r1, c1, r2, c2 in watches:
        assert not (r1 <= 2 <= r2 and c1 <= 1 <= c2)


def test_offset_rebind_drops_stale_watch() -> None:
    def offset_sum(ctx: EvalContext) -> CellValue:
        height = xl_cell(ctx, "S!H1")
        return xl_sum(xl_offset(ctx, ("S", 1, 1), 0, 0, height, 2))

    ctx = _ctx(
        {
            "S!A1": 1.0,
            "S!B1": 2.0,
            "S!A2": 3.0,
            "S!B2": 4.0,
            "S!A3": 5.0,
            "S!B3": 6.0,
            "S!H1": 2.0,
        },
        {"S!Z1": offset_sum},
    )
    assert xl_cell(ctx, "S!Z1") == 10.0
    assert ctx.range_watches["S!Z1"] == [("S", 1, 1, 2, 2)]

    ctx.set_inputs({"S!A3": 50.0})
    assert "S!Z1" in ctx.cache
    assert xl_cell(ctx, "S!Z1") == 10.0

    ctx.set_inputs({"S!H1": 3.0})
    assert xl_cell(ctx, "S!Z1") == 66.0
    assert ctx.range_watches["S!Z1"] == [("S", 1, 1, 3, 2)]

    ctx.set_inputs({"S!B3": 60.0})
    assert xl_cell(ctx, "S!Z1") == 120.0

    ctx.set_inputs({"S!H1": 2.0})
    assert xl_cell(ctx, "S!Z1") == 10.0
    ctx.set_inputs({"S!A3": 999.0})
    assert xl_cell(ctx, "S!Z1") == 10.0


def test_invalidate_clears_helper_cache() -> None:
    def sum_col(ctx: EvalContext) -> CellValue:
        return xl_sum(xl_range(ctx, "S!A1:A2"))

    ctx = _ctx({"S!A1": 1.0, "S!A2": 2.0}, {"S!B1": sum_col})
    xl_cell(ctx, "S!B1")
    ctx.helper_cache[("marker", ())] = 1
    ctx.set_inputs({"S!A1": 5.0})
    assert ctx.helper_cache == {}


def test_slim_context_stays_untracked() -> None:
    ctx = SlimEvalContext(
        inputs=coerce_inputs_dict({"S!A1": 1.0}),
        resolver=lambda _address: None,
    )
    assert not hasattr(ctx, "range_watches")
    assert not hasattr(ctx, "reverse_deps")
    assert not hasattr(ctx, "_record_range_watch")


def test_library_xl_range_records_the_same_watch() -> None:
    from excel_grapher.runtime.cache import xl_range as library_xl_range

    def sum_col(ctx: EvalContext) -> CellValue:
        return xl_sum(library_xl_range(ctx, "S!A1:A3"))

    ctx = _ctx({"S!A1": 1.0, "S!A2": 2.0, "S!A3": 3.0}, {"S!B1": sum_col})
    assert xl_cell(ctx, "S!B1") == 6.0
    assert ctx.range_watches["S!B1"] == [("S", 1, 1, 3, 1)]
    assert "S!A2" not in ctx.reverse_deps
