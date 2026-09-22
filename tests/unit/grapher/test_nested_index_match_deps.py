"""Nested INDEX lookup vectors must not fan out to the whole array (#962)."""

from __future__ import annotations

import tempfile
from pathlib import Path

from fastpyxl import Workbook
from fastpyxl.utils import get_column_letter
from fastpyxl.workbook.defined_name import DefinedName

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    EnumDomain,
    RealIntervalDomain,
    normalize_cell_type_env_key,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefConfig,
    DynamicRefLimits,
    narrow_static_index_lookup_vectors,
)

_REAL = CellType(
    kind=CellKind.NUMBER,
    real_interval=RealIntervalDomain(min=-1e9, max=1e9),
)

_NROWS = 25
_NCOLS = 20


def _string_domain(value: str) -> CellType:
    """Return a singleton string enum domain."""
    return CellType(kind=CellKind.STRING, enum=EnumDomain(values=frozenset({value})))


def _int_domain(value: int) -> CellType:
    """Return a singleton integer enum domain."""
    return CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({value})))


def _config(constraints: dict[str, CellType]) -> DynamicRefConfig:
    """Build a dynamic-ref config from sheet-qualified cell domains."""
    return DynamicRefConfig(
        cell_type_env={
            normalize_cell_type_env_key(addr): cell_type for addr, cell_type in constraints.items()
        },
        limits=DynamicRefLimits(),
    )


def _constraints(nrows: int = _NROWS, ncols: int = _NCOLS) -> dict[str, CellType]:
    constraints: dict[str, CellType] = {
        "Out!C1": _string_domain("CODE_10"),
        "Out!D1": _int_domain(2005),
    }
    for r in range(1, nrows + 2):
        for c in range(1, ncols + 1):
            addr = f"Dump!{get_column_letter(c)}{r}"
            if c == 1 and r >= 2:
                constraints[addr] = _string_domain(f"CODE_{r}")
            elif r == 1 and c > 1:
                constraints[addr] = _int_domain(2000 + c)
            else:
                constraints[addr] = _REAL
    return constraints


def _workbook(formula: str, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Dump"
    for c in range(1, _NCOLS + 1):
        ws.cell(1, c, value=("code" if c == 1 else 2000 + c))
    for r in range(2, _NROWS + 2):
        ws.cell(r, 1, value=f"CODE_{r}")
        for c in range(2, _NCOLS + 1):
            ws.cell(r, c, value=1.0)
    out = wb.create_sheet("Out")
    out["C1"] = "CODE_10"
    out["D1"] = 2005
    last_row = _NROWS + 1
    out["A1"] = formula.format(
        ncols_letter=get_column_letter(_NCOLS),
        last_row=last_row,
    )
    wb.save(path)
    wb.close()


def _dump_deps(formula: str) -> set[str]:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "nested_index_match.xlsx"
        _workbook(formula, path)
        graph = create_dependency_graph(
            path,
            targets=["Out!A1"],
            dynamic_refs=_config(_constraints()),
        )
    return {str(key) for key in graph.get_dependencies("Out!A1") if str(key).startswith("Dump!")}


_VECTOR = (
    "=INDEX(Dump!A1:{ncols_letter}{last_row},"
    "MATCH(C1,Dump!A1:A{last_row},0),"
    "MATCH(D1,Dump!A1:{ncols_letter}1,0))"
)
_NESTED = (
    "=INDEX(Dump!A1:{ncols_letter}{last_row},"
    "MATCH(C1,INDEX(Dump!A1:{ncols_letter}{last_row},,1),0),"
    "MATCH(D1,INDEX(Dump!A1:{ncols_letter}{last_row},1,),0))"
)


def test_narrow_static_index_column_and_row_slices() -> None:
    """Literal INDEX(array,,k) / INDEX(array,k,) become that vector."""
    column = narrow_static_index_lookup_vectors(
        "=MATCH(C1,INDEX(Dump!A1:T26,,1),0)",
        current_sheet="Out",
    )
    row = narrow_static_index_lookup_vectors(
        "=MATCH(D1,INDEX(Dump!A1:T26,1,),0)",
        current_sheet="Out",
    )
    assert column == "=MATCH(C1,Dump!A1:A26,0)"
    assert row == "=MATCH(D1,Dump!A1:T1,0)"


def test_narrow_static_index_keeps_dynamic_selectors_and_quoted_sheets() -> None:
    """Non-literal selectors stay put; quoted sheet names survive the rewrite."""
    dynamic = "=MATCH(C1,INDEX(Dump!A1:T26,,C2),0)"
    assert narrow_static_index_lookup_vectors(dynamic, current_sheet="Out") == dynamic
    quoted = narrow_static_index_lookup_vectors(
        "=MATCH(C1,INDEX('data all'!A1:C4,,1),0)",
        current_sheet="Out",
    )
    assert quoted == "=MATCH(C1,'data all'!A1:A4,0)"
    zero_axis = narrow_static_index_lookup_vectors(
        "=INDEX(Dump!A1:T26,0,2)",
        current_sheet="Dump",
    )
    assert zero_axis == "=Dump!B1:B26"


def test_narrow_nested_static_index_slices_from_the_inside() -> None:
    """INDEX(INDEX(array,,k), row,) resolves to the selected cell of that column."""
    narrowed = narrow_static_index_lookup_vectors(
        "=INDEX(INDEX(Dump!A1:C3,,2),1,)",
        current_sheet="Dump",
    )
    assert narrowed == "=Dump!B1"


def test_nested_index_match_depends_on_lookup_vectors_not_rectangle() -> None:
    """Singleton MATCH needles over INDEX(array,,1) match direct vector MATCH (#962)."""
    vector = _dump_deps(_VECTOR)
    nested = _dump_deps(_NESTED)
    assert nested == vector
    assert "Dump!A10" in nested
    assert "Dump!E10" in nested
    assert "Dump!T1" in nested
    assert "Dump!C5" not in nested
    assert len(nested) < _NROWS * _NCOLS


def test_nested_index_match_over_offset_name_skips_body_cells(tmp_path: Path) -> None:
    """OFFSET-defined table: INDEX(name,,1) depends on the column, not the body."""
    path = tmp_path / "offset_name_index_match.xlsx"
    wb = Workbook()
    data = wb.active
    assert data is not None
    data.title = "data all"
    data["A1"] = "code"
    data["B1"] = 2001
    data["C1"] = 2002
    data["D1"] = 2003
    for row in range(2, 6):
        data.cell(row, 1, value=f"CODE_{row}")
        for col in range(2, 5):
            data.cell(row, col, value=1.0)
    wb.defined_names.add(DefinedName(name="START", attr_text="'data all'!$A$1"))
    wb.defined_names.add(
        DefinedName(
            name="PREV_DSA",
            attr_text="OFFSET(START,0,0,COUNTA('data all'!$A:$A),COUNTA('data all'!$1:$1))",
        )
    )
    out = wb.create_sheet("Out")
    out["C1"] = "CODE_4"
    out["D1"] = 2002
    out["A1"] = "=INDEX(PREV_DSA,MATCH(C1,INDEX(PREV_DSA,,1),0),MATCH(D1,INDEX(PREV_DSA,1,),0))"
    wb.save(path)
    wb.close()

    constraints: dict[str, CellType] = {
        "Out!C1": _string_domain("CODE_4"),
        "Out!D1": _int_domain(2002),
        "'data all'!A1": _REAL,
        "'data all'!B1": _int_domain(2001),
        "'data all'!C1": _int_domain(2002),
        "'data all'!D1": _int_domain(2003),
    }
    for row in range(2, 6):
        constraints[f"'data all'!A{row}"] = _string_domain(f"CODE_{row}")
        constraints[f"'data all'!C{row}"] = _REAL

    graph = create_dependency_graph(
        path,
        targets=["Out!A1"],
        dynamic_refs=_config(constraints),
        capture_dependency_provenance=True,
    )
    deps = {str(key) for key in graph.get_dependencies("Out!A1")}
    assert "'data all'!A4" in deps
    assert "'data all'!C4" in deps
    assert "'data all'!B2" not in deps
    assert "'data all'!D3" not in deps
