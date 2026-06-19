"""Register Excel date/time functions against the shared runtime."""

from __future__ import annotations

from excel_grapher.runtime.datetime import xl_today

from . import register

register("TODAY")(xl_today)
