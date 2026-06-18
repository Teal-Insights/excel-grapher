"""Excel date/time functions for the export runtime."""

from __future__ import annotations

from datetime import date, datetime

from excel_grapher.core.coercions import datetime_to_excel_serial

__all__ = ["xl_today"]


def xl_today() -> float:
    """Return today's date as an Excel serial number."""
    today = datetime.combine(date.today(), datetime.min.time())
    return datetime_to_excel_serial(today)
