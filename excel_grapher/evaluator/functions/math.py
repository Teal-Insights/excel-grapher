"""Register Excel math/stats functions against the shared runtime."""

from __future__ import annotations

from excel_grapher.runtime.math import (
    xl_abs,
    xl_average,
    xl_averageif,
    xl_count,
    xl_counta,
    xl_countif,
    xl_large,
    xl_max,
    xl_min,
    xl_normdist,
    xl_npv,
    xl_rank,
    xl_round,
    xl_rounddown,
    xl_stdev,
    xl_sum,
    xl_sumproduct,
)

from . import register

register("SUM")(xl_sum)
register("AVERAGE")(xl_average)
register("ABS")(xl_abs)
register("MIN")(xl_min)
register("MAX")(xl_max)
register("COUNT")(xl_count)
register("COUNTA")(xl_counta)
register("COUNTIF")(xl_countif)
register("AVERAGEIF")(xl_averageif)
register("SUMPRODUCT")(xl_sumproduct)
register("ROUND")(xl_round)
register("ROUNDDOWN")(xl_rounddown)
register("NPV")(xl_npv)
register("STDEV")(xl_stdev)
register("LARGE")(xl_large)
register("RANK")(xl_rank)
register("NORMDIST")(xl_normdist)
