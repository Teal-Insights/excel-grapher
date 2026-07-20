"""Shared size thresholds for optional operator / SUMPRODUCT acceleration."""

from __future__ import annotations

# Arrays with at least this many cells may materialize for the NumPy fast path.
MIN_OPERATOR_FASTPATH_CELLS = 64
