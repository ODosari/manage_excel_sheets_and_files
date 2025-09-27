from __future__ import annotations

from pathlib import Path
from typing import Iterable


_UTF8_FALLBACKS: tuple[str, ...] = ("utf-8-sig", "utf-8")


def read_text(path: str | Path, encoding: str | None = None) -> str:
    target = Path(path)
    attempts: Iterable[str]
    if encoding:
        attempts = (encoding,)
    else:
        attempts = _UTF8_FALLBACKS
    last_error: Exception | None = None
    for enc in attempts:
        try:
            return target.read_text(encoding=enc)
        except Exception as exc:  # pragma: no cover - bubbled after retries
            last_error = exc
    if last_error:
        raise last_error
    raise ValueError("No encoding attempts were made.")


def write_text(path: str | Path, data: str, *, encoding: str = "utf-8", add_bom: bool = False) -> None:
    target = Path(path)
    text = data
    if add_bom and not text.startswith("\ufeff"):
        text = "\ufeff" + text
    target.write_text(text, encoding=encoding)
