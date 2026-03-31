"""Performance optimization tests for issue #73.

RED phase: these tests fail before the optimizations are applied.

Covers two improvements:
  A. Bounded LRU cache (maxsize=4096) on split_top_level_* and
     _find_function_calls_with_spans — avoids duplicate formula-parse work when
     capture_dependency_provenance=True while keeping memory bounded regardless
     of how many distinct workbooks or formula strings are processed.
  B. Per-BFS-session worksheet cache in create_dependency_graph
     — avoids O(#sheets) fastpyxl.__getitem__ scans on every node visit
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import fastpyxl
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause

# ---------------------------------------------------------------------------
# A. parse-function caching
# ---------------------------------------------------------------------------


def test_split_top_level_if_is_lru_cached() -> None:
    """split_top_level_if must carry a bounded lru_cache so repeat calls are O(1)."""
    from excel_grapher.grapher.parser import split_top_level_if

    # Fails before @lru_cache is added: AttributeError 'function' has no 'cache_info'
    split_top_level_if.cache_clear()

    formula = "=IF(A1>0,B1,C1)"
    split_top_level_if(formula)
    split_top_level_if(formula)  # same formula → must be a cache hit

    info = split_top_level_if.cache_info()
    assert info.hits >= 1, f"Expected at least one cache hit, got {info}"
    assert info.maxsize is not None, "Cache must be bounded (maxsize != None)"


def test_split_top_level_ifs_is_lru_cached() -> None:
    from excel_grapher.grapher.parser import split_top_level_ifs

    split_top_level_ifs.cache_clear()
    formula = "=IFS(A1>0,B1,A1<0,C1)"
    split_top_level_ifs(formula)
    split_top_level_ifs(formula)

    info = split_top_level_ifs.cache_info()
    assert info.hits >= 1, f"Expected cache hit, got {info}"
    assert info.maxsize is not None, "Cache must be bounded"


def test_split_top_level_choose_is_lru_cached() -> None:
    from excel_grapher.grapher.parser import split_top_level_choose

    split_top_level_choose.cache_clear()
    formula = "=CHOOSE(A1,B1,C1)"
    split_top_level_choose(formula)
    split_top_level_choose(formula)

    info = split_top_level_choose.cache_info()
    assert info.hits >= 1
    assert info.maxsize is not None, "Cache must be bounded"


def test_split_top_level_switch_is_lru_cached() -> None:
    from excel_grapher.grapher.parser import split_top_level_switch

    split_top_level_switch.cache_clear()
    formula = "=SWITCH(A1,1,B1,2,C1)"
    split_top_level_switch(formula)
    split_top_level_switch(formula)

    info = split_top_level_switch.cache_info()
    assert info.hits >= 1
    assert info.maxsize is not None, "Cache must be bounded"


def test_find_function_calls_with_spans_is_lru_cached() -> None:
    """_find_function_calls_with_spans must carry lru_cache; fn_names must be frozenset."""
    from excel_grapher.grapher.parser import _find_function_calls_with_spans

    # Fails before @lru_cache (no cache_info) or if fn_names is still set (unhashable)
    _find_function_calls_with_spans.cache_clear()

    formula = "=OFFSET(A1,1,0)+B1"
    fn_names = frozenset({"OFFSET", "INDIRECT"})
    _find_function_calls_with_spans(formula, fn_names)
    _find_function_calls_with_spans(formula, fn_names)

    info = _find_function_calls_with_spans.cache_info()
    assert info.hits >= 1, f"Expected cache hit, got {info}"
    assert info.maxsize is not None, "Cache must be bounded"


# ---------------------------------------------------------------------------
# B. worksheet cache – __getitem__ call-count
# ---------------------------------------------------------------------------


def test_worksheet_cache_reduces_getitem_calls(tmp_path: Path) -> None:
    """wb[sheet] must be called at most once per unique sheet in the BFS loop.

    Before the cache: each of the N formula nodes causes a wb[sheet] call → N calls.
    After the cache : all N nodes share the cached worksheet object → 1 call.
    """
    excel_path = tmp_path / "chain.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    # A1 leaf; A2..A6 formula chain, all on Sheet1 (6 nodes total)
    ws.write_number(0, 0, 1)  # A1 = 1
    ws.write_formula(1, 0, "=A1+1")  # A2
    ws.write_formula(2, 0, "=A2+1")  # A3
    ws.write_formula(3, 0, "=A3+1")  # A4
    ws.write_formula(4, 0, "=A4+1")  # A5
    ws.write_formula(5, 0, "=A5+1")  # A6
    wb.close()

    original_getitem = fastpyxl.Workbook.__getitem__
    calls: list[str] = []

    def spy_getitem(self: fastpyxl.Workbook, key: str):
        calls.append(key)
        return original_getitem(self, key)

    with patch.object(fastpyxl.Workbook, "__getitem__", spy_getitem):
        graph = create_dependency_graph(excel_path, ["Sheet1!A6"], load_values=False)

    assert "Sheet1!A6" in graph  # correctness sanity-check

    sheet1_calls = calls.count("Sheet1")
    # parse_target calls sheetnames (not __getitem__), so BFS loop dominates.
    # With cache: ≤ 1 call per unique sheet name regardless of chain length.
    assert sheet1_calls <= 1, (
        f"Expected wb['Sheet1'] to be called at most once (worksheet cache), "
        f"but it was called {sheet1_calls} times"
    )


# ---------------------------------------------------------------------------
# Correctness: provenance on IF/IFS/CHOOSE/SWITCH with the optimized path
# ---------------------------------------------------------------------------


def test_if_formula_provenance_causes(tmp_path: Path) -> None:
    """Provenance causes must be correct on direct-ref IF formulas."""
    excel_path = tmp_path / "if_prov.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)  # A1 leaf
    ws.write_number(1, 0, 2)  # A2 leaf
    ws.write_number(2, 0, 0)  # A3 condition leaf
    ws.write_formula(3, 0, "=IF(Sheet1!A3,Sheet1!A1,Sheet1!A2)", None, 1)  # A4
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A4"],
        load_values=False,
        capture_dependency_provenance=True,
    )

    for dep in ("Sheet1!A1", "Sheet1!A2", "Sheet1!A3"):
        prov = graph.edge_attrs("Sheet1!A4", dep).get("provenance")
        assert prov is not None, f"Missing provenance for edge A4→{dep}"
        assert DependencyCause.direct_ref in prov.causes, (
            f"Expected direct_ref in causes for A4→{dep}, got {prov.causes}"
        )


def test_choose_formula_provenance_causes(tmp_path: Path) -> None:
    """Provenance causes must be correct on CHOOSE formula."""
    excel_path = tmp_path / "choose_prov.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)  # A1
    ws.write_number(0, 1, 2)  # B1
    ws.write_number(0, 2, 1)  # C1 index
    ws.write_formula(0, 3, "=CHOOSE(Sheet1!C1,Sheet1!A1,Sheet1!B1)", None, 1)  # D1
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!D1"],
        load_values=False,
        capture_dependency_provenance=True,
    )

    for dep in ("Sheet1!A1", "Sheet1!B1", "Sheet1!C1"):
        prov = graph.edge_attrs("Sheet1!D1", dep).get("provenance")
        assert prov is not None, f"Missing provenance for D1→{dep}"
        assert DependencyCause.direct_ref in prov.causes


def test_ifs_formula_provenance_causes(tmp_path: Path) -> None:
    """Provenance causes must be correct on IFS formula."""
    excel_path = tmp_path / "ifs_prov.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 5)   # A1 leaf (value branch)
    ws.write_number(0, 1, 10)  # B1 leaf (value branch)
    ws.write_number(0, 2, 3)   # C1 leaf (condition input)
    # D1 = IFS(C1>5, A1, C1<=5, B1)
    ws.write_formula(0, 3, "=IFS(Sheet1!C1>5,Sheet1!A1,Sheet1!C1<=5,Sheet1!B1)", None, 10)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!D1"],
        load_values=False,
        capture_dependency_provenance=True,
    )

    for dep in ("Sheet1!A1", "Sheet1!B1", "Sheet1!C1"):
        prov = graph.edge_attrs("Sheet1!D1", dep).get("provenance")
        assert prov is not None, f"Missing provenance for D1→{dep}"
        assert DependencyCause.direct_ref in prov.causes, (
            f"Expected direct_ref in causes for D1→{dep}, got {prov.causes}"
        )


def test_switch_formula_provenance_causes(tmp_path: Path) -> None:
    """Provenance causes must be correct on SWITCH formula."""
    excel_path = tmp_path / "switch_prov.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 10)  # A1 result for case 1
    ws.write_number(0, 1, 20)  # B1 result for case 2
    ws.write_number(0, 2, 30)  # C1 default
    ws.write_number(0, 3, 1)   # D1 switch expression input
    # E1 = SWITCH(D1, 1, A1, 2, B1, C1)
    ws.write_formula(
        0, 4, "=SWITCH(Sheet1!D1,1,Sheet1!A1,2,Sheet1!B1,Sheet1!C1)", None, 10
    )
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!E1"],
        load_values=False,
        capture_dependency_provenance=True,
    )

    for dep in ("Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"):
        prov = graph.edge_attrs("Sheet1!E1", dep).get("provenance")
        assert prov is not None, f"Missing provenance for E1→{dep}"
        assert DependencyCause.direct_ref in prov.causes, (
            f"Expected direct_ref in causes for E1→{dep}, got {prov.causes}"
        )
