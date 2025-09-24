import os
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
