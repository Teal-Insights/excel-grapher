from __future__ import annotations

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
    IntervalDomain,
)


def test_cell_type_env_basic_lookup_by_a1_address() -> None:
    env: CellTypeEnv = {
        "Sheet1!B1": CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=0, max=10),
        ),
        "Sheet1!C1": CellType(
            kind=CellKind.NUMBER,
            enum=EnumDomain(values=frozenset({1, 2, 3})),
        ),
    }

    b1 = env["Sheet1!B1"]
    assert b1.kind is CellKind.NUMBER
    assert b1.interval is not None
    assert b1.interval.min == 0
    assert b1.interval.max == 10

    c1 = env["Sheet1!C1"]
    assert c1.kind is CellKind.NUMBER
    assert c1.enum is not None
    assert c1.enum.values == frozenset({1, 2, 3})


def test_cell_type_env_can_mix_kinds_and_domains() -> None:
    env: CellTypeEnv = {
        "Sheet1!A1": CellType(kind=CellKind.STRING),
        "Sheet1!A2": CellType(kind=CellKind.BOOL),
        "Sheet1!A3": CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=-5, max=None),
        ),
    }

    assert env["Sheet1!A1"].kind is CellKind.STRING
    assert env["Sheet1!A1"].interval is None
    assert env["Sheet1!A1"].enum is None

    assert env["Sheet1!A2"].kind is CellKind.BOOL
    assert env["Sheet1!A2"].interval is None
    assert env["Sheet1!A2"].enum is None

    a3 = env["Sheet1!A3"]
    assert a3.kind is CellKind.NUMBER
    assert a3.interval is not None
    assert a3.interval.min == -5
    assert a3.interval.max is None

