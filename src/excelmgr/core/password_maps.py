"""Utilities for loading password maps used to unlock workbooks."""

from __future__ import annotations

import csv
import json
from collections.abc import Mapping
from pathlib import Path

from excelmgr.core.errors import ExcelMgrError


def load_password_map(
    source: str | Mapping[str, str] | None,
    *,
    base_dir: str | Path | None = None,
) -> dict[str, str] | None:
    """Load a password map from JSON/CSV data or a path.

    The ``source`` parameter accepts either a mapping object that is already loaded
    in memory or a filesystem path (absolute or relative) pointing to a JSON or
    CSV document.  When a relative path is provided, ``base_dir`` is used as the
    anchor directory.  The resulting dictionary normalizes keys and values to
    plain strings.
    """

    if source is None:
        return None

    if isinstance(source, Mapping):
        return {str(key): str(value) for key, value in source.items()}

    path = Path(source)
    if not path.is_absolute() and base_dir is not None:
        path = Path(base_dir) / path
    path = path.expanduser().resolve()

    if not path.exists():
        raise ExcelMgrError(f"Password map file not found: {source}")

    if path.suffix.lower() == ".json":
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:  # pragma: no cover - bubble up friendly error
            raise ExcelMgrError(f"Failed to parse password map JSON: {exc}") from exc
        if not isinstance(data, Mapping):
            raise ExcelMgrError("Password map JSON must be an object mapping paths to passwords.")
        return {str(key): str(value) for key, value in data.items()}

    try:
        with open(path, encoding="utf-8", newline="") as fh:
            reader = csv.DictReader(fh)
            if reader.fieldnames is None:
                raise ExcelMgrError("Password CSV must include a header row with 'path' and 'password'.")
            normalized = {name.strip().lower() for name in reader.fieldnames if name}
            required = {"path", "password"}
            if not required.issubset(normalized):
                raise ExcelMgrError("Password CSV must include 'path' and 'password' columns.")
            result: dict[str, str] = {}
            for row in reader:
                lowered = {k.strip().lower(): (v or "") for k, v in row.items() if k}
                key = lowered.get("path", "").strip()
                value = lowered.get("password", "")
                if not key:
                    continue
                result[key] = value
            return result
    except ExcelMgrError:
        raise
    except Exception as exc:  # pragma: no cover - defensive guard
        raise ExcelMgrError(f"Failed to parse password CSV: {exc}") from exc

    raise ExcelMgrError("Unsupported password map format. Use .json or .csv files.")
