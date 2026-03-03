from __future__ import annotations

from typing import Annotated, Literal, TypedDict

from excel_grapher.core.cell_types import (
    Between,
    CellKind,
    CellTypeEnv,
    EnumDomain,
    IntIntervalDomain,
    constraints_to_cell_type_env,
)


class _ConstraintsDict(TypedDict):
    # Integer with an explicit interval domain.
    Sheet1_B1: Annotated[int, Between(0, 10)]
    # Integer literal enum.
    Sheet1_C1: Literal[1, 2, 3]
    # String literal enum, just to show non-numeric enums are supported.
    Sheet1_D1: Literal["NORTH", "SOUTH"]


def test_constraints_mapping_builds_expected_cell_type_env() -> None:
    constraints: _ConstraintsDict = {
        "Sheet1_B1": 5,
        "Sheet1_C1": 2,
        "Sheet1_D1": "NORTH",
    }

    env: CellTypeEnv = constraints_to_cell_type_env(_ConstraintsDict, constraints)

    b1 = env["Sheet1_B1"]
    assert b1.kind is CellKind.NUMBER
    assert b1.interval == IntIntervalDomain(min=0, max=10)
    assert b1.enum is None

    c1 = env["Sheet1_C1"]
    assert c1.kind is CellKind.NUMBER
    assert c1.enum == EnumDomain(values=frozenset({1, 2, 3}))
    assert c1.interval is None

    d1 = env["Sheet1_D1"]
    assert d1.kind is CellKind.STRING
    assert d1.enum == EnumDomain(values=frozenset({"NORTH", "SOUTH"}))
    assert d1.interval is None

