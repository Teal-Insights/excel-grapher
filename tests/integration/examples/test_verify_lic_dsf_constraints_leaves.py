"""LIC-DSF leaf constraint verification over real workbook paths (integration).

Builds temporary workbooks and calls ``verify_lic_dsf_constraints_target_leaves`` so
example tooling rejects invalid target sets before expensive graph exports run.
"""

from __future__ import annotations

from pathlib import Path
from typing import Literal

import fastpyxl
import pytest

from examples.lic_dsf.extract_graph_uncached import (
    collect_lic_dsf_constraint_leaf_violations,
    verify_lic_dsf_constraints_target_leaves,
)


def test_verify_constraints_passes_when_all_targets_are_leaves(tmp_path: Path) -> None:
    path = tmp_path / "wb.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["B1"].value = "text"
    wb.save(path)
    wb.close()

    schema = {
        "Sheet1!A1": Literal[1],
        "Sheet1!B1": Literal["text"],
    }

    verify_lic_dsf_constraints_target_leaves(path, schema)


def test_collect_violations_lists_formula_constrained_cells(tmp_path: Path) -> None:
    path = tmp_path / "wb.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["A2"].value = "=A1+1"
    wb.save(path)
    wb.close()

    schema = {"Sheet1!A2": Literal[99]}

    fc, ms = collect_lic_dsf_constraint_leaf_violations(path, schema)
    assert fc == ["Sheet1!A2"]
    assert ms == []


def test_verify_constraints_raises_when_target_is_formula(tmp_path: Path) -> None:
    path = tmp_path / "wb.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["A2"].value = "=A1+1"
    wb.save(path)
    wb.close()

    schema = {"Sheet1!A2": Literal[99]}

    with pytest.raises(ValueError, match="formulas"):
        verify_lic_dsf_constraints_target_leaves(path, schema)
