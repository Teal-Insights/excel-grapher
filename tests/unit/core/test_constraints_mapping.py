from __future__ import annotations

from typing import Annotated, Literal, TypedDict, cast

from excel_grapher.core.cell_types import (
    Between,
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
    GreaterThanCell,
    IntervalDomain,
    NotEqualCell,
    RealBetween,
    RealIntervalDomain,
    constraints_to_cell_type_env,
)
from excel_grapher.grapher.dynamic_refs import DynamicRefLimits, expand_leaf_env_to_argument_env


# Address-style keys (Sheet1!B1) are the convention for DynamicRefConfig; TypedDict
# cannot use "!" in attribute names, so we set __annotations__ programmatically.
class _ConstraintsDict(TypedDict, total=False):
    pass


_ConstraintsDict.__annotations__["Sheet1!B1"] = Annotated[int, Between(0, 10)]
_ConstraintsDict.__annotations__["Sheet1!C1"] = Literal[1, 2, 3]
_ConstraintsDict.__annotations__["Sheet1!D1"] = Literal["NORTH", "SOUTH"]


def test_constraints_mapping_builds_expected_cell_type_env() -> None:
    constraints = cast(
        _ConstraintsDict,
        {
            "Sheet1!B1": 5,
            "Sheet1!C1": 2,
            "Sheet1!D1": "NORTH",
        },
    )

    env: CellTypeEnv = constraints_to_cell_type_env(_ConstraintsDict, constraints)

    b1 = env["Sheet1!B1"]
    assert b1.kind is CellKind.NUMBER
    assert b1.interval == IntervalDomain(min=0, max=10)
    assert b1.enum is None

    c1 = env["Sheet1!C1"]
    assert c1.kind is CellKind.NUMBER
    assert c1.enum == EnumDomain(values=frozenset({1, 2, 3}))
    assert c1.interval is None

    d1 = env["Sheet1!D1"]
    assert d1.kind is CellKind.STRING
    assert d1.enum == EnumDomain(values=frozenset({"NORTH", "SOUTH"}))
    assert d1.interval is None


class _FloatConstraintsDict(TypedDict, total=False):
    pass


_FloatConstraintsDict.__annotations__["Sheet1!E1"] = Annotated[float, RealBetween(0.0, 1.0)]
_FloatConstraintsDict.__annotations__["Sheet1!F1"] = Annotated[float, RealBetween(-0.5, 0.5)]


def test_constraints_mapping_supports_float_between() -> None:
    constraints = cast(
        _FloatConstraintsDict,
        {
            "Sheet1!E1": 0.5,
            "Sheet1!F1": -0.1,
        },
    )

    env: CellTypeEnv = constraints_to_cell_type_env(_FloatConstraintsDict, constraints)

    e1 = env["Sheet1!E1"]
    assert e1.kind is CellKind.NUMBER
    assert e1.interval is None
    assert e1.real_interval == RealIntervalDomain(min=0.0, max=1.0)
    assert e1.enum is None

    f1 = env["Sheet1!F1"]
    assert f1.kind is CellKind.NUMBER
    assert f1.interval is None
    assert f1.real_interval == RealIntervalDomain(min=-0.5, max=0.5)
    assert f1.enum is None


class _RealIntervalLeafEnvDict(TypedDict, total=False):
    pass


_RealIntervalLeafEnvDict.__annotations__["Sheet1!B1"] = Annotated[float, RealBetween(0.0, 1.0)]


def test_expand_leaf_env_widens_formula_when_dependency_has_real_interval_only() -> None:
    """Real intervals are not enumerable; type expansion falls back to ANY (issue #40)."""
    constraints = cast(_RealIntervalLeafEnvDict, {"Sheet1!B1": 0.5})
    leaf_env = constraints_to_cell_type_env(_RealIntervalLeafEnvDict, constraints)
    out = expand_leaf_env_to_argument_env(
        {"Sheet1!A1"},
        lambda addr: "=Sheet1!B1+1" if addr == "Sheet1!A1" else None,
        lambda _f, _sh: {"Sheet1!B1"},
        leaf_env,
        DynamicRefLimits(),
    )
    assert out["Sheet1!A1"].kind is CellKind.ANY
    assert out["Sheet1!A1"].interval is None
    assert out["Sheet1!A1"].enum is None


class _RelationalConstraintsDict(TypedDict, total=False):
    pass


_RelationalConstraintsDict.__annotations__["Sheet1!A1"] = Annotated[int, Between(0, 10)]
_RelationalConstraintsDict.__annotations__["Sheet1!B1"] = Annotated[
    int,
    Between(1, 20),
    GreaterThanCell("'Sheet1'!A1"),
    NotEqualCell("'Sheet1'!C1"),
]
_RelationalConstraintsDict.__annotations__["Sheet1!C1"] = Annotated[int, Between(0, 20)]


def test_constraints_mapping_preserves_relational_metadata() -> None:
    constraints = cast(
        _RelationalConstraintsDict,
        {
            "Sheet1!A1": 5,
            "Sheet1!B1": 9,
            "Sheet1!C1": 8,
        },
    )

    env: CellTypeEnv = constraints_to_cell_type_env(_RelationalConstraintsDict, constraints)

    b1 = env["Sheet1!B1"]
    assert b1.relations == (
        GreaterThanCell("Sheet1!A1"),
        NotEqualCell("Sheet1!C1"),
    )


class _QuotedSheetConstraintsDict(TypedDict, total=False):
    pass


_SHEET_NEEDS_QUOTES = "Input 4 - External Financing"
_QS_QUOTED_A1 = f"'{_SHEET_NEEDS_QUOTES}'!A1"
_QS_QUOTED_B1 = f"'{_SHEET_NEEDS_QUOTES}'!B1"
_QS_NORMAL_A1 = f"{_SHEET_NEEDS_QUOTES}!A1"
_QS_NORMAL_B1 = f"{_SHEET_NEEDS_QUOTES}!B1"

_QuotedSheetConstraintsDict.__annotations__[_QS_QUOTED_A1] = Annotated[int, Between(0, 10)]
_QuotedSheetConstraintsDict.__annotations__[_QS_QUOTED_B1] = Annotated[
    int,
    Between(1, 20),
    GreaterThanCell(_QS_QUOTED_A1),
]


def test_constraints_mapping_normalizes_quoted_sheet_keys_in_env() -> None:
    """TypedDict keys may use Excel quoting; env keys match _normalize_cell_address (PR #46)."""
    constraints = cast(
        _QuotedSheetConstraintsDict,
        {
            _QS_QUOTED_A1: 1,
            _QS_QUOTED_B1: 5,
        },
    )

    env: CellTypeEnv = constraints_to_cell_type_env(_QuotedSheetConstraintsDict, constraints)

    assert _QS_QUOTED_A1 not in env
    assert _QS_QUOTED_B1 not in env
    assert set(env.keys()) == {_QS_NORMAL_A1, _QS_NORMAL_B1}

    a1 = env[_QS_NORMAL_A1]
    assert a1.kind is CellKind.NUMBER
    assert a1.interval == IntervalDomain(min=0, max=10)

    b1 = env[_QS_NORMAL_B1]
    assert b1.kind is CellKind.NUMBER
    assert b1.relations == (GreaterThanCell(_QS_NORMAL_A1),)


def test_expand_leaf_env_resolves_format_key_addr_against_normalized_env() -> None:
    """Graph builder passes format_key addresses; env keys are normalized (PR #46)."""
    norm = "Chart Data!I21"
    quoted = "'Chart Data'!I21"
    leaf_env: CellTypeEnv = {
        norm: CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=1)),
    }
    out = expand_leaf_env_to_argument_env(
        {quoted},
        lambda _addr: None,
        lambda _f, _sh: set(),
        leaf_env,
        DynamicRefLimits(),
    )
    assert quoted in out
    assert out[quoted].interval == IntervalDomain(min=1, max=1)
