"""Coerce workbook and manifest values into series-binding scalars."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any

from excel_grapher.series_bindings.codegen_literals import py_scalar_literal
from excel_grapher.series_bindings.types import Scalar

_EXCEL_EPOCH = datetime(1899, 12, 30)
_BOOL_TRUE = frozenset({"true", "1", "yes"})
_BOOL_FALSE = frozenset({"false", "0", "no"})

__all__ = ["coerce_constant", "coerce_scalar", "py_scalar_literal"]


def _ensure_naive_datetime(value: datetime) -> datetime:
    if value.tzinfo is None:
        return value
    raise ValueError(f"Timezone-aware datetime values are not supported: {value!r}")


def _normalize_date(value: date) -> datetime:
    return datetime.combine(value, datetime.min.time())


def _parse_iso_datetime(text: str) -> datetime:
    stripped = text.strip()
    if not stripped:
        raise ValueError(f"Cannot coerce {text!r} to datetime")
    if "T" in stripped or " " in stripped:
        parsed = datetime.fromisoformat(stripped.replace("Z", "+00:00"))
        return _ensure_naive_datetime(parsed)
    return _normalize_date(date.fromisoformat(stripped))


def _excel_serial_to_datetime(serial: float) -> datetime:
    days = int(serial)
    fraction = serial - days
    result = _EXCEL_EPOCH + timedelta(days=days)
    if fraction:
        seconds = round(fraction * 86400)
        if seconds:
            result += timedelta(seconds=seconds)
    return result


def _coerce_datetime(raw: Any, *, excel_serial: bool) -> datetime:
    if isinstance(raw, datetime):
        return _ensure_naive_datetime(raw)
    if isinstance(raw, date):
        return _normalize_date(raw)
    if isinstance(raw, str):
        return _parse_iso_datetime(raw)
    if excel_serial and isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return _excel_serial_to_datetime(float(raw))
    raise ValueError(f"Cannot coerce {raw!r} to datetime")


def _coerce_bool(raw: Any) -> bool:
    if isinstance(raw, bool):
        return raw
    if isinstance(raw, (int, float)):
        return bool(raw)
    text = str(raw).strip().lower()
    if text in _BOOL_TRUE:
        return True
    if text in _BOOL_FALSE:
        return False
    raise ValueError(f"Cannot coerce {raw!r} to bool")


def coerce_scalar(raw: Any, read_as: str) -> Scalar:
    """Coerce a workbook or manifest value using an explicit or automatic read mode."""
    if raw is None:
        return None
    if read_as == "auto":
        if isinstance(raw, bool):
            return raw
        if isinstance(raw, int) and not isinstance(raw, bool):
            return raw
        if isinstance(raw, float):
            return raw
        if isinstance(raw, str):
            return raw
        if isinstance(raw, datetime):
            return _ensure_naive_datetime(raw)
        if isinstance(raw, date):
            return _normalize_date(raw)
        return str(raw)
    if read_as == "string":
        return str(raw)
    if read_as == "int":
        return int(raw)
    if read_as == "float":
        return float(raw)
    if read_as == "number":
        if isinstance(raw, int) and not isinstance(raw, bool):
            return raw
        return float(raw)
    if read_as == "bool":
        return _coerce_bool(raw)
    if read_as == "datetime":
        return _coerce_datetime(raw, excel_serial=True)
    raise ValueError(f"Unknown read mode: {read_as!r}")


def coerce_constant(value: Any, *, read_as: str) -> Scalar:
    """Coerce a manifest constant using the effective read mode."""
    return coerce_scalar(value, read_as)
