"""Expand `CellTypeEnv` keys match `_lookup_cell_type` (issue #972)."""

from __future__ import annotations

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnvDict,
    EnumDomain,
    normalize_cell_type_env_key,
)
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefLimits,
    _lookup_cell_type,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
)


class _ScanCountingEnv(CellTypeEnvDict):
    """`CellTypeEnvDict` that counts full-key scans versus membership checks."""

    def __init__(self) -> None:
        super().__init__()
        self.item_scan_ops = 0
        self.contains_ops = 0

    def items(self):
        self.item_scan_ops += 1
        return super().items()

    def keys(self):
        self.item_scan_ops += 1
        return super().keys()

    def __iter__(self):
        self.item_scan_ops += 1
        return super().__iter__()

    def __contains__(self, key: object) -> bool:
        self.contains_ops += 1
        return super().__contains__(key)


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


def test_expand_repeat_on_cell_type_env_dict_does_not_scan_cache() -> None:
    """Repeat expand on a `CellTypeEnvDict` is O(refs), not O(|cache|) (issue #972).

    Builder shares one `CellTypeEnvDict` across INDEX/MATCH sites. Re-canonicalizing
    that store on every call would re-parse every cached key.
    """
    shared = _ScanCountingEnv()
    cell_type = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2014})))
    for row in range(1, 5_001):
        shared[f"Sheet1!A{row}"] = cell_type
    argument_refs = {f"Sheet1!A{row}" for row in range(1, 11)}

    shared.item_scan_ops = 0
    shared.contains_ops = 0
    expand_leaf_env_to_argument_env(
        argument_refs,
        lambda _addr: None,
        lambda _f, _sh: set(),
        shared,
        DynamicRefLimits(),
        shared_cell_type_cache=shared,
    )
    assert shared.item_scan_ops == 0
    assert shared.contains_ops == len(argument_refs)
    assert all(addr in shared for addr in argument_refs)
