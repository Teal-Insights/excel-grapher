from __future__ import annotations

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    CellTypeEnvDict,
    EnumDomain,
    IntervalDomain,
    canonicalize_cell_type_env_keys,
    leaves_missing_cell_type_constraints,
    lookup_cell_type,
    normalize_cell_type_env_key,
)
from excel_grapher.grapher.dynamic_refs import _lookup_cell_type


def test_normalize_cell_type_env_key_strips_excel_sheet_quotes() -> None:
    sheet = "Chart Data"
    assert normalize_cell_type_env_key(f"'{sheet}'!B2") == f"{sheet}!B2"
    assert normalize_cell_type_env_key(f"'{sheet}'!b2") == f"{sheet}!B2"
    assert normalize_cell_type_env_key(f"{sheet}!B2") == f"{sheet}!B2"


def test_lookup_cell_type_reads_normalized_keys_only() -> None:
    """Issue #972: lookups always go through `normalize_cell_type_env_key`."""
    quoted = "'Imported data'!G59"
    norm = normalize_cell_type_env_key(quoted)
    cell_type = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014})))
    assert quoted != norm
    assert lookup_cell_type({quoted: cell_type}, quoted) is None
    assert lookup_cell_type({norm: cell_type}, quoted) is cell_type
    assert _lookup_cell_type({norm: cell_type}, quoted) is cell_type


def test_canonicalize_cell_type_env_keys_rewrites_quoted_keys() -> None:
    quoted = "'Imported data'!G59"
    norm = normalize_cell_type_env_key(quoted)
    cell_type = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014})))
    env = {quoted: cell_type, "Sheet1!A1": cell_type}
    canonicalize_cell_type_env_keys(env)
    assert set(env) == {norm, "Sheet1!A1"}
    assert env[norm] is cell_type


def test_cell_type_env_dict_stores_normalized_keys() -> None:
    quoted = "'Imported data'!G59"
    norm = normalize_cell_type_env_key(quoted)
    cell_type = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014})))
    env = CellTypeEnvDict({quoted: cell_type})
    env["'Imported data'!$g$60"] = CellType(kind=CellKind.NUMBER)

    assert set(env) == {norm, "Imported data!G60"}
    assert env[quoted] is cell_type
    assert env[norm] is cell_type
    assert quoted in env
    assert lookup_cell_type(env, quoted) is cell_type
    assert quoted not in set(env)


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
    assert env.contains_ops <= 20 * 50
