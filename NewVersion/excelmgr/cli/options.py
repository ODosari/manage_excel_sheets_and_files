import csv
import json
import os
from pathlib import Path
from typing import Dict

import typer


def read_secret(password: str | None, password_env: str | None, password_file: str | None) -> str | None:
    if password:
        return password
    if password_env:
        return os.environ.get(password_env)
    if password_file:
        try:
            with open(password_file, "r", encoding="utf-8") as f:
                return f.read().strip()
        except FileNotFoundError:
            raise typer.BadParameter(f"Password file not found: {password_file}")
    return None


def read_password_map(password_map: str | None) -> Dict[str, str] | None:
    if not password_map:
        return None

    path = Path(password_map)
    if not path.exists():
        raise typer.BadParameter(f"Password map file not found: {password_map}")

    try:
        if path.suffix.lower() == ".json":
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise typer.BadParameter("Password map JSON must be an object mapping paths to passwords.")
            result: Dict[str, str] = {}
            for key, value in data.items():
                if not isinstance(key, str) or not isinstance(value, str):
                    raise typer.BadParameter("Password map JSON keys and values must be strings.")
                result[key] = value
            return result

        result = {}
        with open(path, "r", encoding="utf-8", newline="") as fh:
            reader = csv.DictReader(fh)
            required = {"path", "password"}
            if reader.fieldnames is None:
                raise typer.BadParameter("Password CSV must include a header row with 'path' and 'password'.")
            normalized = {name.strip().lower() for name in reader.fieldnames if name}
            if not required.issubset(normalized):
                raise typer.BadParameter("Password CSV must include 'path' and 'password' columns.")
            for row in reader:
                lowered = {k.strip().lower(): (v or "") for k, v in row.items() if k}
                key = lowered.get("path", "").strip()
                value = lowered.get("password", "")
                if not key:
                    continue
                result[key] = value
        return result
    except typer.BadParameter:
        raise
    except Exception as exc:  # pragma: no cover - defensive guard
        raise typer.BadParameter(f"Failed to parse password map: {exc}") from exc
