from __future__ import annotations

import hashlib
import json
from collections.abc import Mapping
from typing import Any


def bindings_canonical_json(bindings: Mapping[str, Any]) -> str:
    """Return stable JSON for cache keys and content hashing."""
    return json.dumps(bindings, sort_keys=True, separators=(",", ":"), ensure_ascii=True)


def bindings_canonical_sha256(bindings: Mapping[str, Any]) -> str:
    """SHA-256 hex digest of the canonical binding manifest."""
    return hashlib.sha256(bindings_canonical_json(bindings).encode("utf-8")).hexdigest()
