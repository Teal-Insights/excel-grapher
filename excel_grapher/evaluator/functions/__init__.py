"""Excel built-in function dispatch registry for the evaluator.

Maps canonical Excel function names to shared runtime callables in
``excel_grapher.runtime``. Implementations live in runtime; this module is
the explicit allowlist for ``FormulaEvaluator`` dispatch.

ROW, COLUMN, COLUMNS, OFFSET, and INDIRECT need range context and are
special-cased in ``excel_grapher.evaluator.evaluator`` instead of appearing here.
"""

from __future__ import annotations

from collections.abc import Callable

from excel_grapher.runtime.datetime import xl_today
from excel_grapher.runtime.info import (
    xl_isblank,
    xl_iserror,
    xl_isna,
    xl_isnumber,
    xl_istext,
    xl_na,
)
from excel_grapher.runtime.logic import xl_and, xl_choose, xl_ifna, xl_not, xl_or
from excel_grapher.runtime.lookup import (
    xl_hlookup,
    xl_index,
    xl_lookup,
    xl_match,
    xl_vlookup,
    xl_xlookup,
)
from excel_grapher.runtime.math import (
    xl_abs,
    xl_average,
    xl_averageif,
    xl_count,
    xl_counta,
    xl_countif,
    xl_exp,
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
from excel_grapher.runtime.reference import xl_address
from excel_grapher.runtime.text import (
    xl_concatenate,
    xl_left,
    xl_lower,
    xl_mid,
    xl_numbervalue,
    xl_right,
    xl_text,
    xl_value,
)

from ..types import CellValue

FUNCTIONS: dict[str, Callable[..., CellValue]] = {
    # datetime
    "TODAY": xl_today,
    # info
    "ISNUMBER": xl_isnumber,
    "ISTEXT": xl_istext,
    "ISBLANK": xl_isblank,
    "ISERROR": xl_iserror,
    "ISNA": xl_isna,
    "NA": xl_na,
    # logic
    "AND": xl_and,
    "OR": xl_or,
    "NOT": xl_not,
    "CHOOSE": xl_choose,
    "IFNA": xl_ifna,
    # lookup
    "INDEX": xl_index,
    "MATCH": xl_match,
    "LOOKUP": xl_lookup,
    "VLOOKUP": xl_vlookup,
    "HLOOKUP": xl_hlookup,
    "XLOOKUP": xl_xlookup,
    # math / stats
    "SUM": xl_sum,
    "AVERAGE": xl_average,
    "ABS": xl_abs,
    "EXP": xl_exp,
    "MIN": xl_min,
    "MAX": xl_max,
    "COUNT": xl_count,
    "COUNTA": xl_counta,
    "COUNTIF": xl_countif,
    "AVERAGEIF": xl_averageif,
    "SUMPRODUCT": xl_sumproduct,
    "ROUND": xl_round,
    "ROUNDDOWN": xl_rounddown,
    "NPV": xl_npv,
    "STDEV": xl_stdev,
    "LARGE": xl_large,
    "RANK": xl_rank,
    "NORMDIST": xl_normdist,
    # reference (ROW / COLUMN / COLUMNS / OFFSET / INDIRECT are special-cased in evaluator)
    "ADDRESS": xl_address,
    # text
    "LEFT": xl_left,
    "RIGHT": xl_right,
    "MID": xl_mid,
    "CONCATENATE": xl_concatenate,
    "TEXT": xl_text,
    "LOWER": xl_lower,
    "VALUE": xl_value,
    "NUMBERVALUE": xl_numbervalue,
}
