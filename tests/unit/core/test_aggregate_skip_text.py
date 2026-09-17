"""#420 / #419 interaction: aggregates vs strict empty-text arithmetic.

Range aggregates (from #419) skip text cells. Literal empty-text args still
go through `to_number` and become `#VALUE!` after #420. SUMPRODUCT treats
text as 0.
"""

from __future__ import annotations

from excel_grapher.core import XlError
from excel_grapher.core.math_funcs import sum_cells
from excel_grapher.core.sumproduct import sumproduct_cells


def test_sum_literal_empty_text_is_value_error() -> None:
    """Literal empty text is coerced via `to_number`, not range-filtered."""
    assert sum_cells(1, "", 2) == XlError.VALUE


def test_sumproduct_treats_text_as_zero() -> None:
    """SUMPRODUCT coerces non-numeric text to 0 (Excel), not `#VALUE!`."""
    assert sumproduct_cells([1, ""], [1, 1]) == 1.0
    assert sumproduct_cells([1, "x"], [2, 3]) == 2.0
