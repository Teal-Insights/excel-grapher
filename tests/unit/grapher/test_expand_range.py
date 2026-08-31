"""Range expansion budget: fail closed over `max_cells`; default cap is 50_000."""

from __future__ import annotations

import inspect

import pytest

from excel_grapher.grapher import DEFAULT_MAX_RANGE_CELLS
from excel_grapher.grapher.builder import (
    create_dependency_graph,
    list_dynamic_ref_constraint_candidates,
)
from excel_grapher.grapher.dynamic_refs import expand_leaf_env_to_argument_env
from excel_grapher.grapher.parser import expand_range
from excel_grapher.grapher.target_expansion import expand_targets_to_roots
from excel_grapher.series_bindings.ranges import (
    effective_reader_range_address,
    expand_data_range,
    expand_data_range_for_graph,
)

DEFAULT_CAP = DEFAULT_MAX_RANGE_CELLS
OVER_DEFAULT = DEFAULT_CAP + 1


def test_expand_range_at_max_cells_enumerates_every_cell() -> None:
    got = expand_range(
        sheet="Sheet1",
        start_col="A",
        start_row=1,
        end_col="A",
        end_row=5000,
        max_cells=5000,
    )
    assert len(got) == 5000
    assert got[0] == ("Sheet1", "A1")
    assert got[-1] == ("Sheet1", "A5000")


def test_expand_range_over_max_cells_raises() -> None:
    with pytest.raises(ValueError, match=r"5001 cells, exceeding max_cells=5000") as exc_info:
        expand_range(
            sheet="Sheet1",
            start_col="A",
            start_row=1,
            end_col="A",
            end_row=5001,
            max_cells=5000,
        )
    assert "Sheet1!A1:A5001" in str(exc_info.value)


def test_expand_range_quoted_sheet_in_error() -> None:
    with pytest.raises(ValueError, match=r"'My Sheet'!A1:C2") as exc_info:
        expand_range(
            sheet="My Sheet",
            start_col="A",
            start_row=1,
            end_col="C",
            end_row=2,
            max_cells=4,
        )
    assert "6 cells" in str(exc_info.value)


def test_expand_range_at_default_cap_enumerates_every_cell() -> None:
    got = expand_range(
        sheet="Sheet1",
        start_col="A",
        start_row=1,
        end_col="A",
        end_row=DEFAULT_CAP,
        max_cells=DEFAULT_CAP,
    )
    assert len(got) == DEFAULT_CAP
    assert got[0] == ("Sheet1", "A1")
    assert got[-1] == ("Sheet1", f"A{DEFAULT_CAP}")


def test_expand_range_over_default_cap_raises() -> None:
    with pytest.raises(
        ValueError, match=rf"{OVER_DEFAULT} cells, exceeding max_cells={DEFAULT_CAP}"
    ) as exc_info:
        expand_range(
            sheet="Sheet1",
            start_col="A",
            start_row=1,
            end_col="A",
            end_row=OVER_DEFAULT,
            max_cells=DEFAULT_CAP,
        )
    assert f"Sheet1!A1:A{OVER_DEFAULT}" in str(exc_info.value)


def test_expand_targets_to_roots_over_max_range_cells_raises() -> None:
    with pytest.raises(ValueError, match="exceeding max_cells=2"):
        expand_targets_to_roots(
            ["Sheet1!A1:A3"],
            sheetnames=["Sheet1"],
            named_ranges={},
            named_range_ranges={},
            max_range_cells=2,
        )


def test_expand_targets_to_roots_default_expands_50k() -> None:
    roots = expand_targets_to_roots(
        ["Sheet1!A1:A50000"],
        sheetnames=["Sheet1"],
        named_ranges={},
        named_range_ranges={},
    )
    assert len(roots) == DEFAULT_CAP
    assert roots[0] == ("Sheet1", "A1")
    assert roots[-1] == ("Sheet1", "A50000")


def test_expand_targets_to_roots_default_overflow_raises() -> None:
    with pytest.raises(ValueError, match=rf"exceeding max_cells={DEFAULT_CAP}"):
        expand_targets_to_roots(
            ["Sheet1!A1:A50001"],
            sheetnames=["Sheet1"],
            named_ranges={},
            named_range_ranges={},
        )


def test_expand_data_range_default_expands_50k() -> None:
    addresses = expand_data_range("Sheet1!A1:A50000")
    assert len(addresses) == DEFAULT_CAP
    assert addresses[0] == "Sheet1!A1"
    assert addresses[-1] == "Sheet1!A50000"


def test_expand_data_range_over_default_cap_raises() -> None:
    with pytest.raises(ValueError, match=rf"exceeding max_cells={DEFAULT_CAP}"):
        expand_data_range("Sheet1!A1:A50001")


def test_max_range_cells_defaults_are_aligned() -> None:
    """Every public default must share the same 50_000-cell budget."""
    assert DEFAULT_MAX_RANGE_CELLS == 50_000
    for fn in (
        create_dependency_graph,
        list_dynamic_ref_constraint_candidates,
        expand_targets_to_roots,
        expand_leaf_env_to_argument_env,
        expand_data_range,
        expand_data_range_for_graph,
        effective_reader_range_address,
    ):
        default = inspect.signature(fn).parameters["max_range_cells"].default
        assert default == DEFAULT_CAP, f"{fn.__qualname__} default is {default}"
