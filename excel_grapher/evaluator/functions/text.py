"""Register Excel text functions against the shared runtime."""

from __future__ import annotations

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

from . import register

register("LEFT")(xl_left)
register("RIGHT")(xl_right)
register("MID")(xl_mid)
register("CONCATENATE")(xl_concatenate)
register("TEXT")(xl_text)
register("LOWER")(xl_lower)
register("VALUE")(xl_value)
register("NUMBERVALUE")(xl_numbervalue)
