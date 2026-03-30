from __future__ import annotations

from pathlib import Path
from typing import Any, Literal, TypedDict

import fastpyxl
import pytest

from example.map_lic_dsf_indicators import (
    collect_lic_dsf_constraint_leaf_violations,
    verify_lic_dsf_constraints_target_leaves,
)


class _LeafOk(TypedDict, total=False):
    pass


class _HasFormula(TypedDict, total=False):
    pass


def test_verify_constraints_passes_when_all_targets_are_leaves(tmp_path: Path) -> None:
    path = tmp_path / "wb.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["B1"].value = "text"
    wb.save(path)
    wb.close()

    constrain(_LeafOk, "Sheet1!A1", Literal[1])
    constrain(_LeafOk, "Sheet1!B1", Literal["text"])

    verify_lic_dsf_constraints_target_leaves(path, _LeafOk)


def test_collect_violations_lists_formula_constrained_cells(tmp_path: Path) -> None:
    path = tmp_path / "wb.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["A2"].value = "=A1+1"
    wb.save(path)
    wb.close()

    constrain(_HasFormula, "Sheet1!A2", Literal[99])

    fc, ms = collect_lic_dsf_constraint_leaf_violations(path, _HasFormula)
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

    constrain(_HasFormula, "Sheet1!A2", Literal[99])

    with pytest.raises(ValueError, match="formulas"):
        verify_lic_dsf_constraints_target_leaves(path, _HasFormula)


def constrain(constraints: type[Any], address: str, annotation: Any) -> None:
    from excel_grapher import constrain as eg_constrain

    eg_constrain(constraints, address, annotation)
