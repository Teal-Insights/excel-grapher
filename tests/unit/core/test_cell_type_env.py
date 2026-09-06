from __future__ import annotations

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
    IntervalDomain,
    leaves_missing_cell_type_constraints,
    normalize_cell_type_env_key,
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


def test_normalize_cell_type_env_key_strips_excel_sheet_quotes() -> None:
    sheet = "Chart Data"
    assert normalize_cell_type_env_key(f"'{sheet}'!B2") == f"{sheet}!B2"
    assert normalize_cell_type_env_key(f"'{sheet}'!b2") == f"{sheet}!B2"
    assert normalize_cell_type_env_key(f"{sheet}!B2") == f"{sheet}!B2"


def test_leaves_missing_cell_type_constraints_ignores_format_key_quoting() -> None:
    env: CellTypeEnv = {
        "Chart Data!I21": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=1)),
    }
    leaves = {"'Chart Data'!I21", "Sheet1!Z9"}
    missing = leaves_missing_cell_type_constraints(leaves, env)
    assert missing == {"Sheet1!Z9"}


class _MembershipCountingEnv(dict[str, CellType]):
    """Dict that counts full-key scans versus O(1) membership checks."""

    def __init__(self, mapping: dict[str, CellType]) -> None:
        super().__init__(mapping)
        self.contains_ops = 0
        self.key_scan_ops = 0

    def __contains__(self, key: object) -> bool:
        self.contains_ops += 1
        return super().__contains__(key)

    def keys(self):
        self.key_scan_ops += 1
        return super().keys()

    def __iter__(self):
        self.key_scan_ops += 1
        return super().__iter__()


def test_leaves_missing_cell_type_constraints_does_not_scan_env_keys() -> None:
    """Issue #715: membership is O(leaves), not an O(|env|) frozenset rebuild."""
    env = _MembershipCountingEnv(
        {f"Sheet1!A{i}": CellType(kind=CellKind.NUMBER) for i in range(1, 5_001)}
    )
    leaves = [f"Sheet1!A{i}" for i in range(1, 51)]
    for _ in range(20):
        missing = leaves_missing_cell_type_constraints(leaves, env)
        assert missing == set()

    assert env.key_scan_ops == 0
    assert env.contains_ops == 20 * 50
