"""Utilities for golden master testing with Excel automation."""

from tests.utils._helpers import is_wsl, wsl_path_to_windows_unc
from tests.utils.modify_and_recalculate import modify_and_recalculate_workbook
from tests.utils.read_cell import read_cell_value, read_range_values

__all__ = [
    "is_wsl",
    "wsl_path_to_windows_unc",
    "read_cell_value",
    "read_range_values",
    "modify_and_recalculate_workbook",
]
