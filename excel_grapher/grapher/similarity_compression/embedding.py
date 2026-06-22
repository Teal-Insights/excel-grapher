"""Embedding providers for similarity-aware compression."""

from __future__ import annotations

import hashlib
import importlib
import math
import struct
from typing import Protocol, runtime_checkable

__all__ = [
    "EmbeddingCache",
    "EmbeddingProvider",
    "MockEmbeddingProvider",
    "OpenAIEmbeddingProvider",
    "embed_texts",
]


@runtime_checkable
class EmbeddingProvider(Protocol):
    """Embed canonical text blobs into dense vectors."""

    def embed(self, texts: list[str]) -> list[list[float]]: ...


class MockEmbeddingProvider:
    """Deterministic hash-based embeddings for tests and offline use."""

    def __init__(self, *, dimensions: int = 32) -> None:
        self._dimensions = dimensions

    def embed(self, texts: list[str]) -> list[list[float]]:
        return [_text_to_unit_vector(text, dimensions=self._dimensions) for text in texts]


class OpenAIEmbeddingProvider:
    """Embed text with the OpenAI embeddings API."""

    def __init__(self, *, model: str = "text-embedding-3-small") -> None:
        self._model = model

    def embed(self, texts: list[str]) -> list[list[float]]:
        if not texts:
            return []
        try:
            openai = importlib.import_module("openai")
        except ImportError as exc:
            raise ImportError(
                "OpenAIEmbeddingProvider requires the optional `embeddings` extra: "
                "uv add --optional embeddings openai"
            ) from exc
        client = openai.OpenAI()
        response = client.embeddings.create(model=self._model, input=texts)
        ordered = sorted(response.data, key=lambda item: item.index)
        return [list(item.embedding) for item in ordered]


class EmbeddingCache:
    """Memoize embeddings by canonical text blob."""

    def __init__(self) -> None:
        self._vectors: dict[str, list[float]] = {}

    def embed(self, texts: list[str], provider: EmbeddingProvider) -> dict[str, list[float]]:
        """Return vectors for ``texts``, reusing cached entries when present."""
        missing = [text for text in texts if text not in self._vectors]
        if missing:
            vectors = provider.embed(missing)
            for text, vector in zip(missing, vectors, strict=True):
                self._vectors[text] = vector
        return {text: self._vectors[text] for text in texts}


def embed_texts(
    texts: list[str],
    provider: EmbeddingProvider,
    *,
    cache: EmbeddingCache | None = None,
) -> dict[str, list[float]]:
    """Embed ``texts`` and return a text-to-vector mapping."""
    if cache is None:
        vectors = provider.embed(texts)
        return dict(zip(texts, vectors, strict=True))
    return cache.embed(texts, provider)


def _text_to_unit_vector(text: str, *, dimensions: int) -> list[float]:
    digest = hashlib.sha256(text.encode("utf-8")).digest()
    values: list[float] = []
    while len(values) < dimensions:
        for offset in range(0, len(digest), 4):
            if len(values) >= dimensions:
                break
            chunk = digest[offset : offset + 4]
            if len(chunk) < 4:
                digest = hashlib.sha256(digest).digest()
                chunk = digest[:4]
            (raw,) = struct.unpack("!I", chunk)
            values.append((raw / 2**32) * 2.0 - 1.0)
        digest = hashlib.sha256(digest).digest()
    norm = math.sqrt(sum(value * value for value in values))
    if norm == 0:
        return values
    return [value / norm for value in values]
