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


# Address-style keys (Sheet1!B1) are the convention for DynamicRefConfig; TypedDict
# cannot use "!" in attribute names, so we set __annotations__ programmatically.
class _ConstraintsDict(TypedDict, total=False):
    pass


_ConstraintsDict.__annotations__["Sheet1!B1"] = Annotated[int, Between(0, 10)]
_ConstraintsDict.__annotations__["Sheet1!C1"] = Literal[1, 2, 3]
_ConstraintsDict.__annotations__["Sheet1!D1"] = Literal["NORTH", "SOUTH"]


def test_constraints_mapping_builds_expected_cell_type_env() -> None:
    constraints: _ConstraintsDict = {
        "Sheet1!B1": 5,
        "Sheet1!C1": 2,
        "Sheet1!D1": "NORTH",
    }

    env: CellTypeEnv = constraints_to_cell_type_env(_ConstraintsDict, constraints)

    b1 = env["Sheet1!B1"]
    assert b1.kind is CellKind.NUMBER
    assert b1.interval == IntIntervalDomain(min=0, max=10)
    assert b1.enum is None

    c1 = env["Sheet1!C1"]
    assert c1.kind is CellKind.NUMBER
    assert c1.enum == EnumDomain(values=frozenset({1, 2, 3}))
    assert c1.interval is None

    d1 = env["Sheet1!D1"]
    assert d1.kind is CellKind.STRING
    assert d1.enum == EnumDomain(values=frozenset({"NORTH", "SOUTH"}))
    assert d1.interval is None

