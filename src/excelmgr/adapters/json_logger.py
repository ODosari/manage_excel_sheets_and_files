"""Structured logging utilities for the CLI."""

from __future__ import annotations

import json
import logging
import sys
import time
import uuid
from typing import Any, TextIO


class JsonLogger:
    """Emit structured log events with a lightweight API."""

    def __init__(
        self,
        *,
        level: int = logging.INFO,
        stream: TextIO | None = None,
        fmt: str = "json",
        file: str | None = None,
    ) -> None:
        self._logger = logging.getLogger("excelmgr")
        self._logger.setLevel(level)
        handler: logging.Handler
        if file is None:
            handler = logging.StreamHandler(stream or sys.stderr)
        else:
            handler = logging.FileHandler(file)
        handler.setFormatter(logging.Formatter("%(message)s"))
        self._logger.handlers = [handler]
        self._run_id = str(uuid.uuid4())
        self._fmt = fmt

    def _serialize(self, payload: dict[str, Any]) -> str:
        if self._fmt == "text":
            keys = sorted(k for k in payload if k not in {"event", "ts", "run_id", "level"})
            parts = " ".join(f"{key}={payload[key]}" for key in keys)
            return f"[{payload['level']}] {payload['event']} {parts}".rstrip()
        return json.dumps(payload)

    def _emit(self, event: str, *, level: str, **kwargs: Any) -> None:
        payload: dict[str, Any] = {
            "event": event,
            "ts": round(time.time(), 3),
            "run_id": self._run_id,
            "level": level,
        }
        payload.update(kwargs)
        level_value = getattr(logging, level.upper(), logging.INFO)
        self._logger.log(level_value, self._serialize(payload))

    def info(self, event: str, **kwargs: Any) -> None:
        self._emit(event, level="INFO", **kwargs)

    def warn(self, event: str, **kwargs: Any) -> None:
        self._emit(event, level="WARNING", **kwargs)

    warning = warn

    def error(self, event: str, **kwargs: Any) -> None:
        self._emit(event, level="ERROR", **kwargs)
