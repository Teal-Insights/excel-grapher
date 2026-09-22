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
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefLimits,
    _lookup_cell_type,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
)


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


def test_expand_leaf_env_keys_match_lookup_normalization() -> None:
    """Expand output keys must match `_lookup_cell_type` (issue #972)."""
    quoted = "'Imported data'!G59"
    norm = normalize_cell_type_env_key(quoted)
    cell_type = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014})))
    out = expand_leaf_env_to_argument_env(
        {quoted},
        lambda _addr: None,
        lambda _f, _sh: set(),
        {norm: cell_type},
        DynamicRefLimits(),
    )
    assert set(out) == {norm}
    assert _lookup_cell_type(out, quoted) == cell_type
    assert _lookup_cell_type(out, norm) == cell_type


def test_expand_leaf_env_infers_quoted_sheet_formula_under_normalized_key() -> None:
    """Pinned formula needles on quoted sheets stay visible after expand."""
    quoted_needle = "'Imported data'!G59"
    quoted_leaf = "'Imported data'!G60"
    norm_needle = normalize_cell_type_env_key(quoted_needle)
    norm_leaf = normalize_cell_type_env_key(quoted_leaf)
    leaf_type = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014})))

    def get_cell_formula(addr: str) -> str | None:
        if normalize_cell_type_env_key(addr) == norm_needle:
            return f"={quoted_leaf}"
        return None

    def get_refs_from_formula(_formula: str, _sheet: str) -> set[str]:
        return {quoted_leaf}

    out = expand_leaf_env_to_argument_env(
        {quoted_needle},
        get_cell_formula,
        get_refs_from_formula,
        {norm_leaf: leaf_type},
        DynamicRefLimits(),
    )
    found = _lookup_cell_type(out, quoted_needle)
    assert set(out) == {norm_needle, norm_leaf}
    assert found is not None
    assert found.enum is not None
    assert found.enum.values == frozenset({2014})


def test_expand_env_feeds_index_match_for_quoted_sheet_needle() -> None:
    """INDEX/MATCH keeps a pinned quoted-sheet needle after expand (issue #972)."""
    sheet = "Imported data"
    quoted_needle = f"'{sheet}'!G59"
    quoted_leaf = f"'{sheet}'!G60"
    norm_needle = normalize_cell_type_env_key(quoted_needle)
    norm_leaf = normalize_cell_type_env_key(quoted_leaf)
    leaves = {
        norm_leaf: CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014}))),
        f"{sheet}!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2013}))),
        f"{sheet}!A2": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014}))),
        f"{sheet}!B1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({10}))),
        f"{sheet}!B2": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({20}))),
    }

    def get_cell_formula(addr: str) -> str | None:
        if normalize_cell_type_env_key(addr) == norm_needle:
            return f"={quoted_leaf}"
        return None

    def get_refs_from_formula(_formula: str, _sheet: str) -> set[str]:
        return {quoted_leaf}

    argument_refs = {
        quoted_needle,
        quoted_leaf,
        f"'{sheet}'!A1",
        f"'{sheet}'!A2",
        f"'{sheet}'!B1",
        f"'{sheet}'!B2",
    }
    out = expand_leaf_env_to_argument_env(
        argument_refs,
        get_cell_formula,
        get_refs_from_formula,
        leaves,
        DynamicRefLimits(),
    )
    formula = f"=INDEX('{sheet}'!B1:B2,MATCH({quoted_needle},'{sheet}'!A1:A2,0),1)"
    targets = infer_dynamic_index_targets(formula, current_sheet=sheet, cell_type_env=out)
    assert {normalize_cell_type_env_key(target) for target in targets} == {f"{sheet}!B2"}


def test_shared_cell_type_cache_hits_across_quoted_and_normalized_refs() -> None:
    """Shared expand cache must treat format_key and env keys as the same cell."""
    quoted = "'Imported data'!G59"
    norm = normalize_cell_type_env_key(quoted)
    cell_type = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014})))
    shared: dict[str, CellType] = {}
    first = expand_leaf_env_to_argument_env(
        {quoted},
        lambda _addr: None,
        lambda _f, _sh: set(),
        {norm: cell_type},
        DynamicRefLimits(),
        shared_cell_type_cache=shared,
    )
    second = expand_leaf_env_to_argument_env(
        {norm},
        lambda _addr: None,
        lambda _f, _sh: set(),
        {norm: cell_type},
        DynamicRefLimits(),
        shared_cell_type_cache=shared,
    )
    assert first is second
    assert set(shared) == {norm}
    assert _lookup_cell_type(shared, quoted) == cell_type


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
