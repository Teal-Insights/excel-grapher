"""Excel function semantics shared by expr_eval, runtime, and export."""

from __future__ import annotations

from .math_funcs import abs_number as xl_abs
from .math_funcs import exp_number as xl_exp

__all__ = ["xl_abs", "xl_exp"]
