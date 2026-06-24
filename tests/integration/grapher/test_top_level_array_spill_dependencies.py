"""Graph extraction expectations for dynamic-array spill side effects."""

from __future__ import annotations

from pathlib import Path
from typing import Literal

import pytest
import xlsxwriter

from excel_grapher import DynamicRefConfig, create_dependency_graph
from tests.fixtures.array_results.workbook import column_compare_path, ensure_committed_fixtures

ensure_committed_fixtures()

ARRAY_ANCHOR = "Data!D10"
STATIC_SPILL_SLOTS = frozenset({"Data!D11", "Data!D12"})


@pytest.mark.xfail(
    strict=True,
    reason=(
        "Issue #284: graph extraction does not infer static array-formula spill "
        "footprints from the AST, so anchor formulas do not depend on occupied "
        "footprint slots for #SPILL! blocking."
    ),
)
def test_static_array_formula_spill_footprint_is_in_anchor_closure() -> None:
    """Static array formulas should pull their spill footprint into the graph."""
    graph = create_dependency_graph(
        column_compare_path(),
        [ARRAY_ANCHOR],
        load_values=True,
    )

    assert set(graph.keys()) >= STATIC_SPILL_SLOTS
    assert graph.get_dependencies(ARRAY_ANCHOR) >= STATIC_SPILL_SLOTS
    for slot in STATIC_SPILL_SLOTS:
        assert ARRAY_ANCHOR in graph.get_dependents(slot)


@pytest.mark.xfail(
    strict=True,
    reason=(
        "Issue #284: graph extraction resolves constrained dynamic refs as input "
        "dependencies but does not use the resolved array shape to add spill "
        "footprint dependencies."
    ),
)
def test_constrained_dynamic_ref_spill_footprint_is_in_anchor_closure(tmp_path: Path) -> None:
    """Constrained dynamic refs should expose a resolvable spill footprint."""
    workbook = _build_constrained_dynamic_spill_workbook(tmp_path / "dynamic_spill.xlsx")
    graph = create_dependency_graph(
        workbook,
        [ARRAY_ANCHOR],
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_constraints({"Data!B1": Literal[3]}, {}),
    )

    assert graph.get_dependencies(ARRAY_ANCHOR) >= {"Data!B1", "Data!C5", "Data!C6", "Data!C7"}
    assert set(graph.keys()) >= STATIC_SPILL_SLOTS
    assert graph.get_dependencies(ARRAY_ANCHOR) >= STATIC_SPILL_SLOTS
    for slot in STATIC_SPILL_SLOTS:
        assert ARRAY_ANCHOR in graph.get_dependents(slot)


def _build_constrained_dynamic_spill_workbook(path: Path) -> Path:
    """Build a workbook whose OFFSET height is fixed by a constrained input cell."""
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Data")
    worksheet.write_number(0, 1, 3)
    for row_offset, category in enumerate(["Software", "Hardware", "Software"]):
        worksheet.write_string(4 + row_offset, 2, category)
    worksheet.write_formula(9, 3, '=OFFSET(Data!C5,0,0,Data!B1,1)="Software"')
    workbook.close()
    return path
