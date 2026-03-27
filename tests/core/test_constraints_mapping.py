from __future__ import annotations

from typing import Annotated, Literal, TypedDict, cast

from excel_grapher.core.cell_types import (
    Between,
    CellKind,
    CellTypeEnv,
    EnumDomain,
    IntervalDomain,
    constraints_to_cell_type_env,
)


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


_FloatConstraintsDict.__annotations__["Sheet1!E1"] = Annotated[float, Between(0.0, 1.0)]
_FloatConstraintsDict.__annotations__["Sheet1!F1"] = Annotated[float, Between(-0.5, 0.5)]


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
    assert e1.interval == IntervalDomain(min=0.0, max=1.0)
    assert e1.enum is None

    f1 = env["Sheet1!F1"]
    assert f1.kind is CellKind.NUMBER
    assert f1.interval == IntervalDomain(min=-0.5, max=0.5)
    assert f1.enum is None

