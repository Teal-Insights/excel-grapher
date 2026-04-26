"""Excel recalculation helper: platform guards and automation entry points (integration).

Asserts ``modify_and_recalculate_workbook`` fails clearly when Excel backends are
unavailable (e.g. Linux CI); full COM/xlwings runs remain opt-in outside this file.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest

from tests.utils.modify_and_recalculate import modify_and_recalculate_workbook


def test_native_linux_without_wsl_raises() -> None:
    """Recalc requires Windows Excel (xlwings/COM); there is no Linux headless fallback."""
    with (
        patch("tests.utils.modify_and_recalculate.is_wsl", return_value=False),
        patch("tests.utils.modify_and_recalculate.sys.platform", "linux"),
        pytest.raises(RuntimeError, match="Excel automation backend"),
    ):
        modify_and_recalculate_workbook(Path("in.xlsm"), Path("out.xlsm"), {})
