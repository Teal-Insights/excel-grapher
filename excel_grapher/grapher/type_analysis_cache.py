"""Persistent memoization cache for dynamic-ref type analysis.

Stores inferred ``CellType`` results for intermediate formula cells in a
SQLite database so that repeated runs can reuse expensive analysis work.
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import logging
import sqlite3
from dataclasses import dataclass
from datetime import UTC, datetime
from importlib.metadata import version as _pkg_version
from pathlib import Path
from typing import Any

from excel_grapher.core.cell_types import (
    CellKind,
    CellRelation,
    CellType,
    CellTypeEnv,
    EnumDomain,
    GreaterThanCell,
    IntervalDomain,
    NotEqualCell,
    RealIntervalDomain,
)
from excel_grapher.grapher.dynamic_refs import DynamicRefLimits

logger = logging.getLogger(__name__)

# Bump this when the analysis logic changes in a way that affects cached results.
ANALYSIS_SCHEMA_VERSION: int = 2

_EXCEL_GRAPHER_VERSION: str = _pkg_version("excel_grapher")

# ---------------------------------------------------------------------------
# CellType JSON serialization
# ---------------------------------------------------------------------------


def _value_to_json(v: Any) -> dict[str, Any]:
    """Serialize a single domain value preserving type distinctions."""
    if isinstance(v, bool):
        return {"t": "bool", "v": v}
    if isinstance(v, int):
        return {"t": "int", "v": v}
    if isinstance(v, float):
        return {"t": "float", "v": v}
    if isinstance(v, str):
        return {"t": "str", "v": v}
    return {"t": "str", "v": str(v)}


def _value_from_json(d: dict[str, Any]) -> Any:
    """Deserialize a single domain value restoring type distinctions."""
    t = d["t"]
    v = d["v"]
    if t == "bool":
        return bool(v)
    if t == "int":
        return int(v)
    if t == "float":
        return float(v)
    if t == "str":
        return str(v)
    return v


def _cell_type_to_json(ct: CellType) -> str:
    """Serialize a ``CellType`` to a stable JSON string."""
    payload: dict[str, Any] = {"kind": ct.kind.value}
    if ct.interval is not None:
        payload["interval"] = {"min": ct.interval.min, "max": ct.interval.max}
    if ct.real_interval is not None:
        payload["real_interval"] = {"min": ct.real_interval.min, "max": ct.real_interval.max}
    if ct.enum is not None:
        serialized = [_value_to_json(v) for v in ct.enum.values]
        serialized.sort(key=lambda x: (x["t"], json.dumps(x["v"], sort_keys=True)))
        payload["enum"] = serialized
    if ct.relations:
        payload["relations"] = [
            {"type": type(r).__name__, "other": r.other}
            for r in sorted(ct.relations, key=lambda r: (type(r).__name__, r.other))
        ]
    return json.dumps(payload, sort_keys=True)


def _cell_type_from_json(s: str) -> CellType:
    """Deserialize a ``CellType`` from a JSON string."""
    d = json.loads(s)
    kind = CellKind(d["kind"])
    interval = None
    if "interval" in d:
        interval = IntervalDomain(min=d["interval"]["min"], max=d["interval"]["max"])
    real_interval = None
    if "real_interval" in d:
        real_interval = RealIntervalDomain(
            min=d["real_interval"]["min"], max=d["real_interval"]["max"]
        )
    enum = None
    if "enum" in d:
        enum = EnumDomain(values=frozenset(_value_from_json(v) for v in d["enum"]))
    relations: tuple[CellRelation, ...] = ()
    if "relations" in d:
        _REL_MAP: dict[str, type[CellRelation]] = {
            "GreaterThanCell": GreaterThanCell,
            "NotEqualCell": NotEqualCell,
        }
        relations = tuple(_REL_MAP[r["type"]](r["other"]) for r in d["relations"])
    return CellType(
        kind=kind, interval=interval, real_interval=real_interval, enum=enum, relations=relations
    )


# ---------------------------------------------------------------------------
# Fingerprinting helpers
# ---------------------------------------------------------------------------


def _compute_limits_fingerprint(limits: DynamicRefLimits) -> str:
    """Produce a deterministic fingerprint for ``DynamicRefLimits``."""
    payload = json.dumps(
        {
            "max_branches": limits.max_branches,
            "max_cells": limits.max_cells,
            "max_depth": limits.max_depth,
        },
        sort_keys=True,
    )
    return hashlib.sha256(payload.encode()).hexdigest()


def _compute_leaf_env_subset_fingerprint(
    consumed_leaf_keys: list[str],
    leaf_env: CellTypeEnv,
) -> str:
    """Hash exactly the consumed leaf entries for deterministic fingerprinting."""
    parts: list[str] = []
    for key in sorted(consumed_leaf_keys):
        ct = leaf_env.get(key)
        if ct is not None:
            parts.append(f"{key}={_cell_type_to_json(ct)}")
        else:
            parts.append(f"{key}=<missing>")
    payload = "\n".join(parts)
    return hashlib.sha256(payload.encode()).hexdigest()


def _compute_key_hash(
    workbook_sha256: str,
    address: str,
    normalized_formula_sha256: str,
    limits_fingerprint: str,
    leaf_env_subset_fingerprint: str,
) -> str:
    """Compute a synthetic primary key hash from all key fields."""
    payload = "\0".join(
        [
            str(ANALYSIS_SCHEMA_VERSION),
            _EXCEL_GRAPHER_VERSION,
            workbook_sha256,
            address,
            normalized_formula_sha256,
            limits_fingerprint,
            leaf_env_subset_fingerprint,
        ]
    )
    return hashlib.sha256(payload.encode()).hexdigest()


def _compute_partial_key_hash(
    workbook_sha256: str,
    address: str,
    normalized_formula_sha256: str,
    limits_fingerprint: str,
) -> str:
    """Compute a lookup key from all fields except leaf-env fingerprint."""
    payload = "\0".join(
        [
            str(ANALYSIS_SCHEMA_VERSION),
            _EXCEL_GRAPHER_VERSION,
            workbook_sha256,
            address,
            normalized_formula_sha256,
            limits_fingerprint,
        ]
    )
    return hashlib.sha256(payload.encode()).hexdigest()


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------


@dataclass
class CacheStats:
    """Lightweight hit/miss/write counters."""

    hits: int = 0
    misses: int = 0
    writes: int = 0
    disabled: bool = False


# ---------------------------------------------------------------------------
# Schema initialization
# ---------------------------------------------------------------------------

_SCHEMA_SQL = """\
CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS formula_cell_types (
    key_hash TEXT PRIMARY KEY,
    partial_key_hash TEXT NOT NULL,
    analysis_schema_version INTEGER NOT NULL,
    excel_grapher_version TEXT NOT NULL,
    workbook_sha256 TEXT NOT NULL,
    address TEXT NOT NULL,
    normalized_formula_sha256 TEXT NOT NULL,
    limits_fingerprint TEXT NOT NULL,
    leaf_env_subset_fingerprint TEXT NOT NULL,
    consumed_leaf_keys_json TEXT NOT NULL,
    cell_type_json TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_formula_cell_types_partial_key
    ON formula_cell_types (partial_key_hash);

CREATE INDEX IF NOT EXISTS idx_formula_cell_types_address
    ON formula_cell_types (address);
"""


# ---------------------------------------------------------------------------
# TypeAnalysisCache
# ---------------------------------------------------------------------------


class TypeAnalysisCache:
    """Persistent cache backed by SQLite for formula-cell type analysis results."""

    def __init__(
        self,
        conn: sqlite3.Connection | None,
        *,
        max_rows: int = 50_000,
        flush_threshold: int = 100,
    ) -> None:
        self._conn = conn
        self._max_rows = max_rows
        self._flush_threshold = flush_threshold
        self._pending: list[dict[str, Any]] = []
        self.stats = CacheStats(disabled=conn is None)

    @classmethod
    def open(
        cls,
        path: str | Path,
        *,
        max_rows: int = 50_000,
        flush_threshold: int = 100,
    ) -> TypeAnalysisCache:
        """Open or create a type-analysis cache at *path*.

        If the file is corrupt or inaccessible, returns a disabled cache that
        silently returns misses and discards writes.
        """
        try:
            path = Path(path)
            path.parent.mkdir(parents=True, exist_ok=True)
            conn = sqlite3.connect(str(path), timeout=5.0)
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA busy_timeout=5000")
            conn.executescript(_SCHEMA_SQL)
            conn.execute(
                "INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)",
                ("schema_version", str(ANALYSIS_SCHEMA_VERSION)),
            )
            conn.execute(
                "INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)",
                ("excel_grapher_version", _EXCEL_GRAPHER_VERSION),
            )
            conn.commit()
            return cls(conn, max_rows=max_rows, flush_threshold=flush_threshold)
        except Exception:
            logger.debug("Failed to open type-analysis cache at %s; disabling", path, exc_info=True)
            return cls(None, max_rows=max_rows, flush_threshold=flush_threshold)

    def close(self) -> None:
        """Flush pending writes and close the database."""
        if self._conn is not None:
            with contextlib.suppress(Exception):
                self.flush()
            with contextlib.suppress(Exception):
                self._conn.close()
        self._conn = None

    def get_formula_cell_type(
        self,
        *,
        workbook_sha256: str,
        address: str,
        normalized_formula_sha256: str,
        limits_fingerprint: str,
        current_leaf_env: CellTypeEnv,
    ) -> tuple[CellType, list[str]] | None:
        """Look up a cached formula-cell type, revalidating the leaf-env fingerprint.

        Queries by partial key (all fields except leaf-env fingerprint), then
        recomputes the leaf-env subset fingerprint from ``current_leaf_env``
        using the stored consumed-leaf keys and compares.

        Returns ``(cell_type, consumed_leaf_keys)`` on hit, or ``None`` on miss.
        """
        if self._conn is None:
            self.stats.misses += 1
            return None
        try:
            partial_hash = _compute_partial_key_hash(
                workbook_sha256,
                address,
                normalized_formula_sha256,
                limits_fingerprint,
            )
            rows = self._conn.execute(
                "SELECT cell_type_json, consumed_leaf_keys_json, leaf_env_subset_fingerprint "
                "FROM formula_cell_types WHERE partial_key_hash = ?",
                (partial_hash,),
            ).fetchall()
            for cell_type_json, stored_keys_json, stored_fp in rows:
                stored_keys: list[str] = json.loads(stored_keys_json)
                recomputed_fp = _compute_leaf_env_subset_fingerprint(stored_keys, current_leaf_env)
                if recomputed_fp == stored_fp:
                    self.stats.hits += 1
                    return _cell_type_from_json(cell_type_json), stored_keys
            self.stats.misses += 1
            return None
        except Exception:
            logger.debug("Cache read error", exc_info=True)
            self.stats.misses += 1
            return None

    def put_formula_cell_type(
        self,
        *,
        workbook_sha256: str,
        address: str,
        normalized_formula_sha256: str,
        limits_fingerprint: str,
        leaf_env_subset_fingerprint: str,
        consumed_leaf_keys: list[str],
        cell_type: CellType,
    ) -> None:
        """Buffer a formula-cell type for later flushing to SQLite."""
        if self._conn is None:
            return
        key_hash = _compute_key_hash(
            workbook_sha256,
            address,
            normalized_formula_sha256,
            limits_fingerprint,
            leaf_env_subset_fingerprint,
        )
        partial_key_hash = _compute_partial_key_hash(
            workbook_sha256,
            address,
            normalized_formula_sha256,
            limits_fingerprint,
        )
        self._pending.append(
            {
                "key_hash": key_hash,
                "partial_key_hash": partial_key_hash,
                "analysis_schema_version": ANALYSIS_SCHEMA_VERSION,
                "excel_grapher_version": _EXCEL_GRAPHER_VERSION,
                "workbook_sha256": workbook_sha256,
                "address": address,
                "normalized_formula_sha256": normalized_formula_sha256,
                "limits_fingerprint": limits_fingerprint,
                "leaf_env_subset_fingerprint": leaf_env_subset_fingerprint,
                "consumed_leaf_keys_json": json.dumps(consumed_leaf_keys),
                "cell_type_json": _cell_type_to_json(cell_type),
                "created_at": datetime.now(UTC).isoformat(),
            }
        )
        if len(self._pending) >= self._flush_threshold:
            self.flush()

    def flush(self) -> None:
        """Write buffered entries to SQLite in a single transaction."""
        if self._conn is None or not self._pending:
            return
        try:
            with self._conn:
                self._conn.executemany(
                    "INSERT OR REPLACE INTO formula_cell_types "
                    "(key_hash, partial_key_hash, analysis_schema_version, "
                    "excel_grapher_version, workbook_sha256, address, "
                    "normalized_formula_sha256, limits_fingerprint, "
                    "leaf_env_subset_fingerprint, consumed_leaf_keys_json, "
                    "cell_type_json, created_at) "
                    "VALUES (:key_hash, :partial_key_hash, :analysis_schema_version, "
                    ":excel_grapher_version, :workbook_sha256, :address, "
                    ":normalized_formula_sha256, :limits_fingerprint, "
                    ":leaf_env_subset_fingerprint, :consumed_leaf_keys_json, "
                    ":cell_type_json, :created_at)",
                    self._pending,
                )
            self.stats.writes += len(self._pending)
            self._pending.clear()
            self._evict_if_needed()
        except Exception:
            logger.debug("Cache flush error", exc_info=True)
            self._pending.clear()

    def _evict_if_needed(self) -> None:
        """Remove oldest rows if count exceeds ``max_rows``."""
        if self._conn is None:
            return
        try:
            row = self._conn.execute("SELECT COUNT(*) FROM formula_cell_types").fetchone()
            if row is None:
                return
            count = row[0]
            if count > self._max_rows:
                excess = count - self._max_rows
                with self._conn:
                    self._conn.execute(
                        "DELETE FROM formula_cell_types WHERE key_hash IN "
                        "(SELECT key_hash FROM formula_cell_types ORDER BY created_at ASC LIMIT ?)",
                        (excess,),
                    )
        except Exception:
            logger.debug("Cache eviction error", exc_info=True)
