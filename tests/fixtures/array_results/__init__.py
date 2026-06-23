"""Workbooks for top-level array formula results (issue #284)."""

from tests.fixtures.array_results.workbook import (
    ARRAY_RESULTS_FIXTURE_DIR,
    build_column_compare_workbook,
    build_numeric_compare_workbook,
    build_row_compare_workbook,
    column_compare_path,
    numeric_compare_path,
    row_compare_path,
)

__all__ = [
    "ARRAY_RESULTS_FIXTURE_DIR",
    "build_column_compare_workbook",
    "build_numeric_compare_workbook",
    "build_row_compare_workbook",
    "column_compare_path",
    "numeric_compare_path",
    "row_compare_path",
]
