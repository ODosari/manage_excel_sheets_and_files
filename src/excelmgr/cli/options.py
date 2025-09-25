import os

import typer

from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.password_maps import load_password_map


def read_secret(password: str | None, password_env: str | None, password_file: str | None) -> str | None:
    if password:
        return password
    if password_env:
        return os.environ.get(password_env)
    if password_file:
        try:
            with open(password_file, encoding="utf-8") as f:
                return f.read().strip()
        except FileNotFoundError:
            raise typer.BadParameter(f"Password file not found: {password_file}") from None
    return None


def read_password_map(password_map: str | None) -> dict[str, str] | None:
    if not password_map:
        return None

    try:
        return load_password_map(password_map)
    except ExcelMgrError as exc:
        raise typer.BadParameter(str(exc)) from exc
