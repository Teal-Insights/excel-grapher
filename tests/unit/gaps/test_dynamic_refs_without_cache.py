"""Dynamic-ref resolution gaps without ``use_cached_dynamic_refs``."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import DynamicRefError, create_dependency_graph
from tests.unit.gaps.workbook_helpers import write_index_match_best_week, write_text_index_match


def test_index_match_graph_build_requires_cached_values(tmp_path: Path) -> None:
    """Dynamic ``INDEX``/``MATCH`` needs cached resolution by default."""
    workbook = write_index_match_best_week(tmp_path / "best_week.xlsx")
    with pytest.raises(DynamicRefError, match="INDEX"):
        create_dependency_graph(
            workbook,
            ["Aggregate Stats!D28"],
            load_values=True,
            use_cached_dynamic_refs=False,
        )


def test_text_index_match_graph_build_requires_cached_values(tmp_path: Path) -> None:
    """``TEXT(INDEX(...,MATCH(MAX(...))))`` (financial_model ``B22`` shape) needs cache."""
    workbook = write_text_index_match(tmp_path / "revenue_summary.xlsx")
    with pytest.raises(DynamicRefError, match="INDEX"):
        create_dependency_graph(
            workbook,
            ["Revenue Model!B22"],
            load_values=True,
            use_cached_dynamic_refs=False,
        )
