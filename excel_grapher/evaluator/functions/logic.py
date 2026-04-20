"""Register Excel logical functions against the shared runtime."""

from __future__ import annotations

from excel_grapher.runtime.logic import xl_and, xl_choose, xl_ifna, xl_or

from . import register

register("AND")(xl_and)
register("OR")(xl_or)
register("CHOOSE")(xl_choose)
register("IFNA")(xl_ifna)
register("_XLFN.IFNA")(xl_ifna)
