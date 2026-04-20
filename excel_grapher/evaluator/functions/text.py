"""Register Excel text functions against the shared runtime."""

from __future__ import annotations

from excel_grapher.runtime.text import (
    xl__xlfn_numbervalue,
    xl_concatenate,
    xl_left,
    xl_mid,
    xl_right,
    xl_text,
)

from . import register

register("LEFT")(xl_left)
register("RIGHT")(xl_right)
register("MID")(xl_mid)
register("CONCATENATE")(xl_concatenate)
register("TEXT")(xl_text)
register("NUMBERVALUE")(xl__xlfn_numbervalue)
register("_XLFN.NUMBERVALUE")(xl__xlfn_numbervalue)
