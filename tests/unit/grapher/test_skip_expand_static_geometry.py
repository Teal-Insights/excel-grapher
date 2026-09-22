"""Skip argument-env expansion when INDEX/MATCH uses static-array bounds (#757)."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated
from unittest.mock import patch

import pytest
from fastpyxl import Workbook
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    EnumDomain,
    RealBetween,
    normalize_cell_type_env_key,
)
from excel_grapher.core.formula_ast import parse as parse_ast
from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    DynamicRefLimits,
    DynamicRefTraceEvent,
    dynamic_ref_selectors_boundable_without_expand,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
    trace_dynamic_refs,
)

Bounded = Annotated[float, RealBetween(-1e9, 1e9)]


def test_index_match_static_lookup_is_boundable_without_env() -> None:
    formula = "=INDEX(Out!B2:V12,MATCH(1,Out!A2:A12,0),MATCH(Out!B1,Out!B1:V1,0))"
    ast = parse_ast(formula)
    assert dynamic_ref_selectors_boundable_without_expand(
        formula,
        current_sheet="Out",
        ast=ast,
    )


def test_index_cell_selector_is_not_boundable_without_env() -> None:
    formula = "=INDEX(Sheet1!A1:A3,Sheet1!B1,1)"
    assert not dynamic_ref_selectors_boundable_without_expand(
        formula,
        current_sheet="Sheet1",
    )


def test_indirect_is_not_boundable_without_env() -> None:
    formula = '=INDIRECT("Sheet1!A1")'
    assert not dynamic_ref_selectors_boundable_without_expand(
        formula,
        current_sheet="Sheet1",
    )


def test_offset_match_is_boundable_without_env() -> None:
    formula = "=OFFSET(Chart!A1,MATCH(Inputs!A1,Lookup!A2:A8,0)-1,0)"
    assert dynamic_ref_selectors_boundable_without_expand(
        formula,
        current_sheet="Outputs",
    )


def test_offset_cell_height_is_not_boundable_without_env() -> None:
    formula = "=OFFSET(Sheet1!A1,0,0,Sheet1!B1,1)"
    assert not dynamic_ref_selectors_boundable_without_expand(
        formula,
        current_sheet="Sheet1",
    )


def test_omitted_index_axis_match_is_boundable_without_env() -> None:
    """INDEX(array,,k) / INDEX(array,k,) inside MATCH are geometry (#968)."""
    formula = (
        "=INDEX('data all'!A2:BH239,"
        "MATCH('Imported data'!D62,INDEX('data all'!A2:BH239,,2),0),"
        "MATCH('Imported data'!K59,INDEX('data all'!A2:BH239,1,),0))"
    )
    assert dynamic_ref_selectors_boundable_without_expand(
        formula,
        current_sheet="Imported data",
        limits=DynamicRefLimits(),
        current_row=62,
        current_col=11,
    )
    whole_column = (
        "=INDEX('data all'!A2:C4,MATCH('Imported data'!A1,INDEX('data all'!A2:C4,0,1),0),1)"
    )
    assert dynamic_ref_selectors_boundable_without_expand(
        whole_column,
        current_sheet="Imported data",
    )


def test_nested_static_index_array_is_boundable_without_env() -> None:
    """INDEX(INDEX(array,,k), MATCH) narrows to a static column before the check."""
    formula = "=INDEX(INDEX(Data!A1:C3,,2),MATCH(Data!E1,Data!A1:A3,0))"
    assert dynamic_ref_selectors_boundable_without_expand(
        formula,
        current_sheet="Data",
    )


def test_omitted_index_axis_without_densifiable_sibling_is_not_boundable() -> None:
    """An omitted axis still needs the other selector to be known from geometry."""
    assert not dynamic_ref_selectors_boundable_without_expand(
        "=INDEX('data all'!A1:C10,,'Imported data'!D1)",
        current_sheet="Imported data",
    )
    assert not dynamic_ref_selectors_boundable_without_expand(
        "=INDEX('data all'!A1:C10,,)",
        current_sheet="Imported data",
    )


def _build_index_match_closure_workbook(path: Path, *, n: int) -> None:
    """INDEX/MATCH whose MATCH lookup cells SUM a large nested-IF block.

    Matches the issue #757 MCVE: expand would type-infer `n` Engine formulas,
    while abstract INDEX only needs the 11x21 Out array geometry.
    """
    wb = Workbook()
    engine = wb.active
    engine.title = "Engine"
    engine["A1"] = 1.0
    engine["B1"] = 2.0
    engine["C1"] = 3.0
    engine["D1"] = 4.0
    for r in range(2, n + 2):
        engine[f"G{r}"] = "=IF(A1>0,IF(B1>0,A1+B1,C1),IF(C1>0,D1,A1))"
    out = wb.create_sheet("Out")
    for c, year in enumerate(range(2020, 2042), start=2):
        out.cell(1, c, year)
    for r in range(2, 13):
        out[f"A{r}"] = f"=IF(SUM(Engine!G2:G{n + 1})>0,1,0)"
        for c in range(2, 23):
            out.cell(r, c, float(r * 10 + c))
    out["Z1"] = "=INDEX(B2:V12,MATCH(1,A2:A12,0),MATCH(B1,B1:V1,0))"
    wb.save(path)


def _index_match_constraints() -> dict[str, object]:
    constraints: dict[str, object] = {
        "Engine!A1": Bounded,
        "Engine!B1": Bounded,
        "Engine!C1": Bounded,
        "Engine!D1": Bounded,
    }
    for c in range(2, 23):
        constraints[f"Out!{get_column_letter(c)}1"] = Bounded
    for r in range(2, 13):
        for c in range(2, 23):
            constraints[f"Out!{get_column_letter(c)}{r}"] = Bounded
    return constraints


def test_index_match_skips_expand_and_argument_subgraph(tmp_path: Path) -> None:
    """MATCH lookup formula closures must not be type-walked for abstract INDEX."""
    n = 80
    excel_path = tmp_path / "index-match-skip.xlsx"
    _build_index_match_closure_workbook(excel_path, n=n)
    constraints = _index_match_constraints()
    config = DynamicRefConfig.from_constraints(constraints)

    expand_calls = 0
    subgraph_calls = 0
    original_expand = expand_leaf_env_to_argument_env

    def counting_expand(*args: object, **kwargs: object):
        nonlocal expand_calls
        expand_calls += 1
        return original_expand(*args, **kwargs)

    from excel_grapher.grapher.dynamic_ref_walk import DynamicRefWalkContext

    original_subgraph = DynamicRefWalkContext.argument_subgraph_refs

    def counting_subgraph(self: object, argument_addrs: object):
        nonlocal subgraph_calls
        subgraph_calls += 1
        return original_subgraph(self, argument_addrs)

    events: list[DynamicRefTraceEvent] = []
    with (
        patch(
            "excel_grapher.grapher.builder.expand_leaf_env_to_argument_env",
            side_effect=counting_expand,
        ),
        patch(
            "excel_grapher.grapher.provenance_collect.expand_leaf_env_to_argument_env",
            side_effect=counting_expand,
        ),
        patch.object(DynamicRefWalkContext, "argument_subgraph_refs", counting_subgraph),
        trace_dynamic_refs(events.append),
    ):
        graph = create_dependency_graph(
            excel_path,
            ["Out!Z1"],
            load_values=True,
            dynamic_refs=config,
            capture_dependency_provenance=True,
        )

    assert expand_calls == 0, f"expected no expand, got {expand_calls}"
    assert subgraph_calls == 0, f"expected no argument-subgraph walk, got {subgraph_calls}"
    assert any(e.kind == "expand-env-skipped" for e in events)
    abstract = [e for e in events if e.kind == "index-abstract"]
    assert abstract, f"expected index-abstract, got {[e.kind for e in events]}"
    assert abstract[0].detail.get("targets") == 11 * 21
    assert "Out!B2" in graph.get_dependencies("Out!Z1")
    assert "Out!V12" in graph.get_dependencies("Out!Z1")
    assert "Out!A2" in graph.get_dependencies("Out!Z1")
    assert "Engine!G2" not in graph.get_dependencies("Out!Z1")


def test_index_match_does_not_require_lookup_formula_leaf_constraints(
    tmp_path: Path,
) -> None:
    """Abstract INDEX/MATCH must succeed without constraints on MATCH lookup leaves."""
    excel_path = tmp_path / "index-match-unconstrained-lookup.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A2"] = 1
    ws["A3"] = 0
    ws["A4"] = 0
    ws["B2"] = 10
    ws["B3"] = 20
    ws["B4"] = 30
    ws["C1"] = "=INDEX(B2:B4,MATCH(1,A2:A4,0))"
    wb.save(excel_path)

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!C1"],
        load_values=False,
        dynamic_refs=DynamicRefConfig(cell_type_env={}, limits=DynamicRefLimits()),
    )
    assert "Sheet1!B2" in graph.get_dependencies("Sheet1!C1")
    assert "Sheet1!B4" in graph.get_dependencies("Sheet1!C1")


def test_quoted_sheet_omitted_index_match_skips_expand(tmp_path: Path) -> None:
    """Omitted-axis INDEX/MATCH keeps the lazy env and stays under `max_cells` (#968).

    The 6x4 grid exceeds `max_cells`. Year headers are pinned; code cells are
    not, so the row MATCH stays unbounded. The untyped corner stays a year
    candidate and the certain 2018 hit stops the scan, so targets are columns
    A and C. Argument-env expansion still does not run.
    """
    excel_path = tmp_path / "quoted-index-match.xlsx"
    wb = Workbook()
    data = wb.active
    assert data is not None
    data.title = "data all"
    data["A1"] = "code"
    data["B1"] = 2017
    data["C1"] = 2018
    data["D1"] = 2019
    for row in range(2, 7):
        data.cell(row, 1, f"KEY.{row}")
        for col in range(2, 5):
            data.cell(row, col, float(row * col))
    imported = wb.create_sheet("Imported data")
    imported["A1"] = "KEY.3"
    imported["B1"] = 2018
    imported["C1"] = (
        "=INDEX('data all'!A1:D6,"
        "MATCH(A1,INDEX('data all'!A1:D6,,1),0),"
        "MATCH(B1,INDEX('data all'!A1:D6,1,),0))"
    )
    wb.save(excel_path)

    def _pinned_number(value: int) -> CellType:
        return CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({value})))

    env = {
        normalize_cell_type_env_key("'Imported data'!B1"): _pinned_number(2018),
        normalize_cell_type_env_key("'data all'!B1"): _pinned_number(2017),
        normalize_cell_type_env_key("'data all'!C1"): _pinned_number(2018),
        normalize_cell_type_env_key("'data all'!D1"): _pinned_number(2019),
    }
    limits = DynamicRefLimits(max_cells=20)
    expand_calls = 0
    original_expand = expand_leaf_env_to_argument_env

    def counting_expand(*args: object, **kwargs: object):
        nonlocal expand_calls
        expand_calls += 1
        return original_expand(*args, **kwargs)

    with (
        patch(
            "excel_grapher.grapher.builder.expand_leaf_env_to_argument_env",
            side_effect=counting_expand,
        ),
        patch(
            "excel_grapher.grapher.provenance_collect.expand_leaf_env_to_argument_env",
            side_effect=counting_expand,
        ),
    ):
        graph = create_dependency_graph(
            excel_path,
            ["'Imported data'!C1"],
            load_values=True,
            dynamic_refs=DynamicRefConfig(cell_type_env=env, limits=limits),
        )

    deps = set(graph.get_dependencies("'Imported data'!C1"))
    assert expand_calls == 0
    assert "'data all'!A3" in deps
    assert "'data all'!C3" in deps
    assert "'data all'!B2" not in deps
    assert "'data all'!D2" not in deps


def test_quoted_sheet_blank_corner_is_numeric_zero_without_expand(tmp_path: Path) -> None:
    """A blank_ranges corner is numeric 0 during exact MATCH, without expand (#979).

    Year headers are pinned and the corner is absent from the cell-type env.
    Needle 2018 misses that 0, so targets are only the 2018 column. Needle 0
    is a certain hit on the corner. Argument-env expansion still does not run.
    """
    excel_path = tmp_path / "quoted-blank-corner.xlsx"
    wb = Workbook()
    data = wb.active
    assert data is not None
    data.title = "data all"
    data["A1"] = "code"
    data["B1"] = 2017
    data["C1"] = 2018
    data["D1"] = 2019
    for row in range(2, 7):
        data.cell(row, 1, f"KEY.{row}")
        for col in range(2, 5):
            data.cell(row, col, float(row * col))
    imported = wb.create_sheet("Imported data")
    imported["A1"] = "KEY.3"
    imported["B1"] = 2018
    imported["C1"] = (
        "=INDEX('data all'!A1:D6,"
        "MATCH(A1,INDEX('data all'!A1:D6,,1),0),"
        "MATCH(B1,INDEX('data all'!A1:D6,1,),0))"
    )
    wb.save(excel_path)

    def _pinned_number(value: int) -> CellType:
        return CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({value})))

    limits = DynamicRefLimits(max_cells=20)
    blanks = ("'data all'!A1",)
    expand_calls = 0
    original_expand = expand_leaf_env_to_argument_env
    original_index = infer_dynamic_index_targets
    captured_targets: list[set[str]] = []

    def counting_expand(*args: object, **kwargs: object):
        nonlocal expand_calls
        expand_calls += 1
        return original_expand(*args, **kwargs)

    def capturing_index(*args: object, **kwargs: object) -> set[str]:
        targets = original_index(*args, **kwargs)
        captured_targets.append(set(targets))
        return targets

    def _targets_for(needle: int) -> set[str]:
        env = {
            normalize_cell_type_env_key("'Imported data'!B1"): _pinned_number(needle),
            normalize_cell_type_env_key("'data all'!B1"): _pinned_number(2017),
            normalize_cell_type_env_key("'data all'!C1"): _pinned_number(2018),
            normalize_cell_type_env_key("'data all'!D1"): _pinned_number(2019),
        }
        captured_targets.clear()
        with (
            patch(
                "excel_grapher.grapher.builder.expand_leaf_env_to_argument_env",
                side_effect=counting_expand,
            ),
            patch(
                "excel_grapher.grapher.provenance_collect.expand_leaf_env_to_argument_env",
                side_effect=counting_expand,
            ),
            patch(
                "excel_grapher.grapher.builder.infer_dynamic_index_targets",
                side_effect=capturing_index,
            ),
        ):
            create_dependency_graph(
                excel_path,
                ["'Imported data'!C1"],
                load_values=True,
                blank_ranges=blanks,
                dynamic_refs=DynamicRefConfig(cell_type_env=env, limits=limits),
            )
        assert len(captured_targets) == 1
        return captured_targets[0]

    year_targets = _targets_for(2018)
    assert expand_calls == 0
    assert year_targets == {f"'data all'!C{row}" for row in range(1, 7)}

    zero_targets = _targets_for(0)
    assert expand_calls == 0
    assert "'data all'!A1" in zero_targets
    assert zero_targets == {f"'data all'!A{row}" for row in range(1, 7)}


def test_index_cell_selector_still_fail_closes_without_constraint(tmp_path: Path) -> None:
    excel_path = tmp_path / "index-cell-selector.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = 30
    ws["B1"] = 2
    ws["C1"] = "=INDEX(A1:A3,B1)"
    wb.save(excel_path)

    with pytest.raises(DynamicRefError, match="no constraint"):
        create_dependency_graph(
            excel_path,
            ["Sheet1!C1"],
            load_values=False,
            dynamic_refs=DynamicRefConfig(cell_type_env={}, limits=DynamicRefLimits()),
        )
