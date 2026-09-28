"""Formula inference of text-or-number union domains (issue #1005)."""

from __future__ import annotations

import pytest

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    EnumDomain,
    IntervalDomain,
    RealIntervalDomain,
    is_union_domain,
)
from excel_grapher.grapher import dynamic_refs as dynamic_refs_mod
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefError,
    DynamicRefLimits,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
)

_REAL = CellType(kind=CellKind.NUMBER, real_interval=RealIntervalDomain(min=-1.0, max=1.0))


def _infer(formula: str, leaf: CellType, *, max_branches: int = 8) -> CellType:
    """Type `Sheet1!B1` holding `formula` over the single leaf `Sheet1!A1`."""
    env = expand_leaf_env_to_argument_env(
        {"Sheet1!B1"},
        lambda addr: formula if addr == "Sheet1!B1" else None,
        lambda _formula, _sheet: {"Sheet1!A1"},
        {"Sheet1!A1": leaf},
        DynamicRefLimits(max_branches=max_branches),
    )
    return env["Sheet1!B1"]


def _int_enum(*values: int) -> CellType:
    return CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset(values)))


def _members(ct: CellType) -> set[object]:
    """Finite membership of a union whose numeric arm is an integer interval."""
    assert ct.enum is not None
    assert ct.interval is not None and ct.interval.min is not None and ct.interval.max is not None
    return set(ct.enum.values) | set(range(ct.interval.min, ct.interval.max + 1))


def test_if_int_enum_or_sentinel_infers_union() -> None:
    ct = _infer('=IF(Sheet1!A1>0,Sheet1!A1,"n.a.")', _int_enum(0, 1, 2))
    assert ct.kind is CellKind.ANY
    assert is_union_domain(ct)
    assert _members(ct) == {1, 2, "n.a."}


def test_if_sparse_int_enum_stays_exact() -> None:
    ct = _infer('=IF(Sheet1!A1>0,Sheet1!A1,"n.a.")', _int_enum(-4, 1, 5))
    assert is_union_domain(ct)
    assert _members(ct) == {1, 5, "n.a."}


def test_if_real_interval_or_sentinel_infers_union() -> None:
    ct = _infer('=IF(Sheet1!A1>0,Sheet1!A1,"n.a.")', _REAL)
    assert ct == CellType(
        kind=CellKind.ANY,
        enum=EnumDomain(values=frozenset({"n.a."})),
        real_interval=RealIntervalDomain(min=-1.0, max=1.0),
    )


def test_isnumber_guard_over_authored_union_keeps_union() -> None:
    authored = CellType(
        kind=CellKind.ANY,
        enum=EnumDomain(values=frozenset({"n.a."})),
        real_interval=RealIntervalDomain(min=0.0, max=2.5),
    )
    ct = _infer('=IF(ISNUMBER(Sheet1!A1),Sheet1!A1,"n.a.")', authored)
    assert ct == authored


def test_iferror_real_or_sentinel_infers_union() -> None:
    ct = _infer('=IFERROR(Sheet1!A1,"n.a.")', _REAL)
    assert ct.enum == EnumDomain(values=frozenset({"n.a."}))
    assert ct.real_interval == RealIntervalDomain(min=-1.0, max=1.0)


def test_non_integral_float_literal_becomes_real_arm() -> None:
    ct = _infer('=IF(Sheet1!A1>0,0.5,"n.a.")', _REAL)
    assert ct.enum is not None and "n.a." in ct.enum.values
    assert is_union_domain(ct)
    assert ct.interval is None


def test_wide_int_interval_union_does_not_enumerate() -> None:
    """A span over `max_branches` used to fail fallback enumeration."""
    wide = CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=10**6))
    ct = _infer('=IF(Sheet1!A1>0,Sheet1!A1,"n.a.")', wide)
    assert ct == CellType(
        kind=CellKind.ANY,
        enum=EnumDomain(values=frozenset({"n.a."})),
        interval=IntervalDomain(min=1, max=10**6),
    )


def test_text_only_branches_over_real_precedent_infer_string_enum() -> None:
    ct = _infer('=IF(Sheet1!A1>0,"pos","neg")', _REAL)
    assert ct == CellType(kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"pos", "neg"})))


def test_error_only_branches_stay_bare_any() -> None:
    ct = _infer("=IF(Sheet1!A1>0,#N/A,#DIV/0!)", _REAL)
    assert ct == CellType(kind=CellKind.ANY)


def test_uncovered_branch_stays_bare_any() -> None:
    ct = _infer('=IF(Sheet1!A1>0,Sheet1!A1*2,"n.a.")', _REAL)
    assert ct == CellType(kind=CellKind.ANY)


def test_fallback_enumeration_parses_formula_once(monkeypatch: pytest.MonkeyPatch) -> None:
    calls = 0
    real_parse = dynamic_refs_mod.parse_ast

    def counting_parse(text: str):
        nonlocal calls
        calls += 1
        return real_parse(text)

    monkeypatch.setattr(dynamic_refs_mod, "parse_ast", counting_parse)
    # Text concatenation is outside numeric analysis, so the fallback enumerates.
    ct = _infer('=Sheet1!A1&"x"', _int_enum(1, 22, 333, 4444))
    assert ct.enum == EnumDomain(values=frozenset({"1x", "22x", "333x", "4444x"}))
    assert calls <= 2


def _match_env(dump_a1: CellType, needle: CellType) -> dict[str, CellType]:
    return {
        "calc!A1": needle,
        "dump!A1": dump_a1,
        "dump!A2": needle,
    }


_MATCH_FORMULA = "=INDEX(dump!B1:B2,MATCH(calc!A1,dump!A1:A2,0),1)"


def test_exact_match_inferred_union_keeps_number_inside_interval() -> None:
    inferred = _infer('=IF(Sheet1!A1>0,Sheet1!A1,"n.a.")', _REAL)
    env = _match_env(inferred, _int_enum(0))
    targets = infer_dynamic_index_targets(_MATCH_FORMULA, current_sheet="calc", cell_type_env=env)
    assert targets == {"dump!B1", "dump!B2"}


def test_exact_match_inferred_union_misses_unrelated_text() -> None:
    inferred = _infer('=IF(Sheet1!A1>0,Sheet1!A1,"n.a.")', _REAL)
    nope = CellType(kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"nope"})))
    env = _match_env(inferred, nope)
    targets = infer_dynamic_index_targets(_MATCH_FORMULA, current_sheet="calc", cell_type_env=env)
    assert targets == {"dump!B2"}


@pytest.mark.parametrize("leaf", [_REAL, _int_enum(0, 1, 2)])
def test_enumerating_inferred_text_union_fails_closed(leaf: CellType) -> None:
    inferred = _infer('=IF(Sheet1!A1>0,Sheet1!A1,"n.a.")', leaf)
    with pytest.raises(DynamicRefError, match="union domain"):
        dynamic_refs_mod._build_domains(["X!A1"], {"X!A1": inferred}, DynamicRefLimits())


def test_offset_over_inferred_text_union_fails_closed() -> None:
    env = expand_leaf_env_to_argument_env(
        {"Sheet1!B1"},
        lambda addr: '=IF(Sheet1!A1>0,Sheet1!A1,"n.a.")' if addr == "Sheet1!B1" else None,
        lambda _formula, _sheet: {"Sheet1!A1"},
        {"Sheet1!A1": _int_enum(0, 1, 2)},
        DynamicRefLimits(),
    )
    with pytest.raises(DynamicRefError, match="union domain"):
        dynamic_refs_mod.infer_dynamic_offset_targets(
            "=OFFSET(Sheet1!C1,Sheet1!B1,0)", current_sheet="Sheet1", cell_type_env=env
        )


def test_fallback_does_not_drop_interval_arm_of_union_precedent() -> None:
    """Enumerating only the enum arm would claim every result is "n.a.x"."""
    authored = CellType(
        kind=CellKind.ANY,
        enum=EnumDomain(values=frozenset({"n.a."})),
        real_interval=RealIntervalDomain(min=0.0, max=2.5),
    )
    assert _infer('=Sheet1!A1&"x"', authored) == CellType(kind=CellKind.ANY)


def test_fallback_enumerates_finite_integer_union_precedent() -> None:
    authored = CellType(
        kind=CellKind.ANY,
        enum=EnumDomain(values=frozenset({7})),
        interval=IntervalDomain(min=1, max=2),
    )
    ct = _infer('=Sheet1!A1&"x"', authored)
    assert ct.enum == EnumDomain(values=frozenset({"1x", "2x", "7x"}))
