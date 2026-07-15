"""LIC-DSF CI-score chain: arithmetic OFFSET named ranges must not collapse to 1x1.

Regression for issue #410: `DSF__DATA_{CPIA,REMGDP,GDPUSD}` use
`COUNTA(...)+n` extents. When those collapsed to the anchor cell, INDEX/MATCH
returned `#REF!` into CI Summary remittances / GDP-USD / score inputs.
"""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest
from fastpyxl.utils.cell import coordinate_from_string

from excel_grapher import FormulaEvaluator, XlError, create_dependency_graph
from excel_grapher.grapher.resolver import build_named_range_map
from tests.utils.excel_workbook_parity import assert_workbook_parity

WORKBOOK_PATH = Path("examples/lic_dsf/data/lic-dsf-template-2025-08-12.xlsm")

# Direct consumers of the arithmetic-OFFSET data tables + CI Summary mirrors.
_CI_CHAIN_TARGETS: list[str] = [
    "Imported data!L75",  # INDEX(DSF__DATA_REMGDP, ...)
    "Imported data!L76",  # INDEX(DSF__DATA_GDPUSD, ...)
    "Imported data!K82",  # IFERROR(INDEX(DSF__DATA_CPIA, ...), ...)
    "CI Summary!C56",  # remittances mirror of L75
    "CI Summary!C57",  # GDP-USD mirror of L76
    "CI Summary!H50",  # remittance ratio using remittances/GDP
]

RTOL = 1e-5
ATOL = 1e-9

_ARITH_OFFSET_NAMES: tuple[str, ...] = (
    "DSF__DATA_CPIA",
    "DSF__DATA_GDPUSD",
    "DSF__DATA_REMGDP",
)


@pytest.mark.slow
def test_lic_dsf_arith_offset_names_resolve_to_multicell_tables() -> None:
    """COUNTA(...)+n defined names must expand past the A1:A1 collapse."""
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")

    wb = fastpyxl.load_workbook(WORKBOOK_PATH, data_only=False, read_only=True, keep_vba=True)
    try:
        maps = build_named_range_map(wb)
    finally:
        wb.close()

    for name in _ARITH_OFFSET_NAMES:
        assert name in maps.range_map, f"{name} missing from range_map"
        sheet, start, end = maps.range_map[name]
        assert start == "A1", f"{name}: unexpected start {start}"
        assert end != "A1", f"{name} collapsed to 1x1 anchor on {sheet} ({start}:{end})"
        # These padded COUNTA+n tables always span multiple columns (headers + years).
        end_col, _end_row = coordinate_from_string(end)
        assert end_col != "A", f"{name}: expected multi-column end, got {end}"


@pytest.mark.slow
def test_lic_dsf_ci_chain_evaluator_matches_excel_cache() -> None:
    """Imported-data INDEX rows and CI Summary mirrors match Excel cached values."""
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")

    graph = create_dependency_graph(
        WORKBOOK_PATH,
        _CI_CHAIN_TARGETS,
        load_values=True,
        max_depth=50,
        use_cached_dynamic_refs=True,
    )
    assert_workbook_parity(
        graph,
        _CI_CHAIN_TARGETS,
        rtol=RTOL,
        atol=ATOL,
        fail_fast=True,
    )


@pytest.mark.slow
def test_lic_dsf_ci_index_formulas_not_substituted_as_a1_a1() -> None:
    """Normalized INDEX formulas should reference the full padded data tables."""
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")

    graph = create_dependency_graph(
        WORKBOOK_PATH,
        ["Imported data!L75", "Imported data!L76", "Imported data!K82"],
        load_values=True,
        max_depth=30,
        use_cached_dynamic_refs=True,
    )
    for addr, needle in (
        ("Imported data!L75", "data_REM_FINAL!A1:A1"),
        ("Imported data!L76", "data_GDPUSD!A1:A1"),
        ("Imported data!K82", "CPIA!A1:A1"),
    ):
        node = graph.get_node(addr)
        assert node is not None
        assert node.normalized_formula is not None
        assert needle not in node.normalized_formula, (
            f"{addr} still references collapsed {needle}: {node.normalized_formula}"
        )

    with FormulaEvaluator(graph) as ev:
        for addr in ("Imported data!L75", "Imported data!L76"):
            result = ev._evaluate_cell(addr)
            assert not isinstance(result, XlError), f"{addr} returned {result}"
