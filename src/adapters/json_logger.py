import json, logging, sys, time, uuid
from typing import Any, Dict, Optional

class JsonLogger:
    def __init__(self, level: int = logging.INFO, stream=None, fmt: str = "json", file: Optional[str] = None):
        self.log = logging.getLogger("excelmgr")
        self.log.setLevel(level)
        handler = logging.StreamHandler(stream or sys.stderr) if file is None else logging.FileHandler(file)
        handler.setFormatter(logging.Formatter("%(message)s"))
        self.log.handlers = [handler]
        self.run_id = str(uuid.uuid4())
        self.fmt = fmt

    def _serialize(self, payload: Dict[str, Any]) -> str:
        if self.fmt == "text":
            items = [f"{k}={payload[k]}" for k in sorted(payload.keys()) if k not in ("event","ts","run_id","level")]
            return f"[{payload['level']}] {payload['event']} " + " ".join(items)
        return json.dumps(payload)

    def _emit(self, event: str, **kwargs: Any):
        payload: Dict[str, Any] = {"event": event, "ts": round(time.time(),3), "run_id": self.run_id}
        payload.update(kwargs)
        level_name = payload.setdefault("level","INFO")
        level = getattr(logging, level_name.upper(), logging.INFO)
        self.log.log(level, self._serialize(payload))

    def info(self, event: str, **kwargs: Any): self._emit(event, level="INFO", **kwargs)
    def warn(self, event: str, **kwargs: Any): self._emit(event, level="WARNING", **kwargs)
    warning = warn
    def error(self, event: str, **kwargs: Any): self._emit(event, level="ERROR", **kwargs)
