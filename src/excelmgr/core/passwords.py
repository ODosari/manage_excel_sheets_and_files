from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path


def resolve_password(path: str, default: str | None, mapping: Mapping[str, str] | None) -> str | None:
    if not mapping:
        return default

    p = Path(path)
    candidates = [str(path), str(p), str(p.resolve()), p.name]
    seen: set[str] = set()
    for candidate in candidates:
        if candidate in seen:
            continue
        seen.add(candidate)
        if candidate in mapping:
            return mapping[candidate]
    return default
