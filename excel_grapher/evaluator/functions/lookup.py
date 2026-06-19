"""Register Excel lookup functions against the shared runtime."""

from __future__ import annotations

from excel_grapher.runtime.lookup import (
    xl_hlookup,
    xl_index,
    xl_lookup,
    xl_match,
    xl_vlookup,
    xl_xlookup,
)

from . import register

register("INDEX")(xl_index)
register("MATCH")(xl_match)
register("LOOKUP")(xl_lookup)
register("VLOOKUP")(xl_vlookup)
register("HLOOKUP")(xl_hlookup)
register("XLOOKUP")(xl_xlookup)
