"""Integration tests: ``create_dependency_graph`` accepts mixed target forms.

Targets may be sheet-qualified single cells, sheet-qualified ranges, or named
ranges (single-cell or rectangular). Each form must seed BFS with the union of
concrete sheet-qualified single-cell roots, deduplicated across overlapping
inputs.
"""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest
from fastpyxl.workbook.defined_name import DefinedName

from excel_grapher import create_dependency_graph


def _build_grid_workbook(path: Path) -> None:
    """Workbook with a small numeric grid plus a defined name covering it.

    - Sheet1!B1..B3 are integer leaves.
    - Sheet1!C1..C3 are formulas referencing the corresponding B cell.
    - 'My Sheet'!A1..B2 form a quoted-sheet rectangle of integer leaves.
    - Defined names: ``OneCell`` -> Sheet1!$C$1, ``BeeCol`` -> Sheet1!$B$1:$B$3.
    """
    wb = fastpyxl.Workbook()
    s1 = wb.active
    s1.title = "Sheet1"
    s1["B1"].value = 1
    s1["B2"].value = 2
    s1["B3"].value = 3
    s1["C1"].value = "=B1+1"
    s1["C2"].value = "=B2+1"
    s1["C3"].value = "=B3+1"

    s2 = wb.create_sheet("My Sheet")
    s2["A1"].value = 10
    s2["A2"].value = 20
    s2["B1"].value = 30
    s2["B2"].value = 40

    wb.defined_names.add(DefinedName("OneCell", attr_text="Sheet1!$C$1"))
    wb.defined_names.add(DefinedName("BeeCol", attr_text="Sheet1!$B$1:$B$3"))

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def test_targets_accept_sheet_qualified_range(tmp_path: Path) -> None:
    """A sheet-qualified range target expands to one root per cell."""
    excel_path = tmp_path / "range_target.xlsx"
    _build_grid_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!C1:C3"], load_values=False)

    for key in ("Sheet1!C1", "Sheet1!C2", "Sheet1!C3"):
        assert key in graph, f"expected expanded root {key!r} in graph"

    assert graph.get_dependencies("Sheet1!C1") == {"Sheet1!B1"}
    assert graph.get_dependencies("Sheet1!C2") == {"Sheet1!B2"}
    assert graph.get_dependencies("Sheet1!C3") == {"Sheet1!B3"}


def test_targets_accept_quoted_sheet_range(tmp_path: Path) -> None:
    """Quoted sheet names with embedded space must work in range targets."""
    excel_path = tmp_path / "quoted_range_target.xlsx"
    _build_grid_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["'My Sheet'!A1:B2"], load_values=False)

    for key in (
        "'My Sheet'!A1",
        "'My Sheet'!A2",
        "'My Sheet'!B1",
        "'My Sheet'!B2",
    ):
        assert key in graph, f"expected expanded root {key!r} in graph"


def test_targets_accept_named_range_single_cell(tmp_path: Path) -> None:
    """A bare defined name pointing to a single cell resolves to that cell."""
    excel_path = tmp_path / "named_cell_target.xlsx"
    _build_grid_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["OneCell"], load_values=False)

    assert "Sheet1!C1" in graph
    assert graph.get_dependencies("Sheet1!C1") == {"Sheet1!B1"}


def test_targets_accept_named_range_rectangle(tmp_path: Path) -> None:
    """A bare defined name pointing to a rectangle expands to all cells."""
    excel_path = tmp_path / "named_range_target.xlsx"
    _build_grid_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["BeeCol"], load_values=False)

    for key in ("Sheet1!B1", "Sheet1!B2", "Sheet1!B3"):
        assert key in graph, f"expected expanded root {key!r} in graph"


def test_targets_accept_mixed_cells_ranges_and_named_ranges(
    tmp_path: Path,
) -> None:
    """One call accepts every supported target form simultaneously."""
    excel_path = tmp_path / "mixed_targets.xlsx"
    _build_grid_workbook(excel_path)

    graph = create_dependency_graph(
        excel_path,
        [
            "Sheet1!C2",
            "Sheet1!C1:C3",
            "'My Sheet'!A1:B2",
            "OneCell",
            "BeeCol",
        ],
        load_values=False,
    )

    for key in (
        "Sheet1!C1",
        "Sheet1!C2",
        "Sheet1!C3",
        "Sheet1!B1",
        "Sheet1!B2",
        "Sheet1!B3",
        "'My Sheet'!A1",
        "'My Sheet'!A2",
        "'My Sheet'!B1",
        "'My Sheet'!B2",
    ):
        assert key in graph, f"expected expanded root {key!r} in graph"


def test_target_nodes_are_marked_with_is_target_metadata(tmp_path: Path) -> None:
    """Expanded target roots should be marked target-aware on graph nodes."""
    excel_path = tmp_path / "target_metadata.xlsx"
    _build_grid_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!C2"], load_values=False)
    target_node = graph.get_node("Sheet1!C2")
    assert target_node is not None
    assert target_node.is_target is True

    dependency_only = graph.get_node("Sheet1!B2")
    assert dependency_only is not None
    assert dependency_only.is_target is False


def test_targets_deduplicate_overlaps(tmp_path: Path) -> None:
    """Overlapping target inputs do not duplicate nodes."""
    excel_path = tmp_path / "overlap_targets.xlsx"
    _build_grid_workbook(excel_path)

    graph_single = create_dependency_graph(excel_path, ["Sheet1!C1:C3"], load_values=False)
    graph_overlap = create_dependency_graph(
        excel_path,
        ["Sheet1!C1:C3", "Sheet1!C2", "OneCell"],
        load_values=False,
    )

    assert sorted(graph_overlap) == sorted(graph_single)


def test_targets_unknown_name_raises_clear_error(tmp_path: Path) -> None:
    """A bare token that is neither sheet-qualified nor a defined name errors."""
    excel_path = tmp_path / "unknown_name.xlsx"
    _build_grid_workbook(excel_path)

    with pytest.raises(ValueError) as exc:
        create_dependency_graph(excel_path, ["NoSuchName"], load_values=False)
    msg = str(exc.value)
    assert "NoSuchName" in msg


def test_targets_reject_invalid_sheet_in_range(tmp_path: Path) -> None:
    """Range target on a missing sheet raises a sheet-not-found ValueError."""
    excel_path = tmp_path / "missing_sheet_range.xlsx"
    _build_grid_workbook(excel_path)

    with pytest.raises(ValueError, match="Sheet not found"):
        create_dependency_graph(excel_path, ["NoSheet!A1:B2"], load_values=False)


def test_targets_accept_sheet_qualified_endpoint_range(tmp_path: Path) -> None:
    """``Sheet!A1:Sheet!B2`` form (normalized Excel range) is accepted."""
    excel_path = tmp_path / "qualified_endpoint.xlsx"
    _build_grid_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!C1:Sheet1!C3"], load_values=False)
    for key in ("Sheet1!C1", "Sheet1!C2", "Sheet1!C3"):
        assert key in graph


def test_targets_reject_empty_string(tmp_path: Path) -> None:
    """An empty target string raises a clear ValueError."""
    excel_path = tmp_path / "empty_target.xlsx"
    _build_grid_workbook(excel_path)

    with pytest.raises(ValueError):
        create_dependency_graph(excel_path, [""], load_values=False)
