"""Register Excel info-type predicates against the shared runtime."""

from __future__ import annotations

from excel_grapher.runtime.info import (
    xl_isblank,
    xl_iserror,
    xl_isna,
    xl_isnumber,
    xl_istext,
    xl_na,
)

from . import register

register("ISNUMBER")(xl_isnumber)
register("ISTEXT")(xl_istext)
register("ISBLANK")(xl_isblank)
register("ISERROR")(xl_iserror)
register("ISNA")(xl_isna)
register("NA")(xl_na)
