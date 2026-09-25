from __future__ import annotations

from typing import Annotated, Any, Literal

import pytest

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
    normalize_cell_type_env_key,
)
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefLimits,
    _lookup_cell_type,
    expand_leaf_env_to_argument_env,
)


def test_constraints_mapping_builds_expected_cell_type_env() -> None:
    schema: dict[str, Any] = {
        "Sheet1!B1": Annotated[int, Between(0, 10)],
        "Sheet1!C1": Literal[1, 2, 3],
        "Sheet1!D1": Literal["NORTH", "SOUTH"],
    }
    constraints = {
        "Sheet1!B1": 5,
        "Sheet1!C1": 2,
        "Sheet1!D1": "NORTH",
    }

    env: CellTypeEnv = constraints_to_cell_type_env(schema, constraints)

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


def test_constraints_mapping_unions_literal_with_real_interval() -> None:
    schema: dict[str, Any] = {
        "Sheet1!A1": Literal["n.a."] | Annotated[float, RealBetween(-1.0e15, 1.0e15)],
    }
    env: CellTypeEnv = constraints_to_cell_type_env(schema, {"Sheet1!A1": "n.a."})
    cell = env["Sheet1!A1"]
    assert cell.kind is CellKind.ANY
    assert cell.enum == EnumDomain(values=frozenset({"n.a."}))
    assert cell.real_interval == RealIntervalDomain(min=-1.0e15, max=1.0e15)
    assert cell.interval is None


def test_integer_union_enumerates_enum_and_interval() -> None:
    from excel_grapher.grapher.dynamic_refs import _build_domains, _build_value_domains

    schema: dict[str, Any] = {
        "Sheet1!A1": Literal[-1, 0, 1] | Annotated[int, Between(0, 3)],
    }
    env = constraints_to_cell_type_env(schema, {})
    assert _build_domains(["Sheet1!A1"], env, DynamicRefLimits()) == {"Sheet1!A1": [-1, 0, 1, 2, 3]}
    assert _build_value_domains(["Sheet1!A1"], env, DynamicRefLimits()) == {
        "Sheet1!A1": [-1, 0, 1, 2, 3]
    }


def test_text_sentinel_union_is_not_enumerated() -> None:
    from excel_grapher.grapher.dynamic_refs import DynamicRefError, _build_domains

    schema: dict[str, Any] = {
        "Sheet1!A1": Literal["n.a."] | Annotated[float, RealBetween(-1.0, 1.0)],
    }
    env = constraints_to_cell_type_env(schema, {})
    with pytest.raises(DynamicRefError, match="union"):
        _build_domains(["Sheet1!A1"], env, DynamicRefLimits())


def test_constraints_mapping_merges_literal_union_arms() -> None:
    env = constraints_to_cell_type_env(
        {"Sheet1!A1": Literal["a"] | Literal["b"]},
        {},
    )
    cell = env["Sheet1!A1"]
    assert cell.enum == EnumDomain(values=frozenset({"a", "b"}))
    assert cell.interval is None
    assert cell.real_interval is None


def test_constraints_mapping_rejects_open_numeric_arm() -> None:
    with pytest.raises(ValueError, match="unconstrained"):
        constraints_to_cell_type_env(
            {"Sheet1!A1": Literal["n.a."] | float},
            {},
        )


def test_constraints_mapping_rejects_two_interval_arms() -> None:
    schema: dict[str, Any] = {
        "Sheet1!A1": Annotated[int, Between(0, 1)] | Annotated[float, RealBetween(0.0, 1.0)],
    }
    with pytest.raises(ValueError, match="between and real_between"):
        constraints_to_cell_type_env(schema, {})


def test_constraints_mapping_supports_float_between() -> None:
    schema: dict[str, Any] = {
        "Sheet1!E1": Annotated[float, RealBetween(0.0, 1.0)],
        "Sheet1!F1": Annotated[float, RealBetween(-0.5, 0.5)],
    }
    constraints = {
        "Sheet1!E1": 0.5,
        "Sheet1!F1": -0.1,
    }

    env: CellTypeEnv = constraints_to_cell_type_env(schema, constraints)

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


def test_expand_leaf_env_widens_formula_when_dependency_has_real_interval_only() -> None:
    """Real intervals are not enumerable; type expansion falls back to ANY (issue #40)."""
    schema: dict[str, Any] = {"Sheet1!B1": Annotated[float, RealBetween(0.0, 1.0)]}
    constraints = {"Sheet1!B1": 0.5}
    leaf_env = constraints_to_cell_type_env(schema, constraints)
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


def test_constraints_mapping_preserves_relational_metadata() -> None:
    schema: dict[str, Any] = {
        "Sheet1!A1": Annotated[int, Between(0, 10)],
        "Sheet1!B1": Annotated[
            int,
            Between(1, 20),
            GreaterThanCell("'Sheet1'!A1"),
            NotEqualCell("'Sheet1'!C1"),
        ],
        "Sheet1!C1": Annotated[int, Between(0, 20)],
    }
    constraints = {
        "Sheet1!A1": 5,
        "Sheet1!B1": 9,
        "Sheet1!C1": 8,
    }

    env: CellTypeEnv = constraints_to_cell_type_env(schema, constraints)

    b1 = env["Sheet1!B1"]
    assert b1.relations == (
        GreaterThanCell("Sheet1!A1"),
        NotEqualCell("Sheet1!C1"),
    )


_SHEET_NEEDS_QUOTES = "Input 4 - External Financing"
_QS_QUOTED_A1 = f"'{_SHEET_NEEDS_QUOTES}'!A1"
_QS_QUOTED_B1 = f"'{_SHEET_NEEDS_QUOTES}'!B1"
_QS_NORMAL_A1 = f"{_SHEET_NEEDS_QUOTES}!A1"
_QS_NORMAL_B1 = f"{_SHEET_NEEDS_QUOTES}!B1"


def test_constraints_mapping_normalizes_quoted_sheet_keys_in_env() -> None:
    """Schema keys may use Excel quoting; env keys match normalized addresses (PR #46)."""
    schema: dict[str, Any] = {
        _QS_QUOTED_A1: Annotated[int, Between(0, 10)],
        _QS_QUOTED_B1: Annotated[
            int,
            Between(1, 20),
            GreaterThanCell(_QS_QUOTED_A1),
        ],
    }
    constraints = {
        _QS_QUOTED_A1: 1,
        _QS_QUOTED_B1: 5,
    }

    env: CellTypeEnv = constraints_to_cell_type_env(schema, constraints)

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
    assert quoted != norm
    assert normalize_cell_type_env_key(quoted) == norm
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
    assert set(out) == {norm}
    assert quoted not in set(out)
    assert out[norm].interval == IntervalDomain(min=1, max=1)
    assert _lookup_cell_type(out, quoted) == out[norm]
    assert _lookup_cell_type(out, norm) == out[norm]
