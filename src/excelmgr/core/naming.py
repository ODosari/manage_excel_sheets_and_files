import re
from typing import Final

_MAX_SHEET: Final[int] = 31
_ILLEGAL_PATTERN: Final[re.Pattern[str]] = re.compile(r"[\\/:?*\[\]]")
_CTRL_PATTERN: Final[re.Pattern[str]] = re.compile(r"[\x00-\x1F]")


def sanitize_sheet_name(name: str) -> str:
    n = str(name).strip()
    if not n:
        n = "Empty"
    # Replace control characters and Excel-illegal characters with safe underscores
    n = _CTRL_PATTERN.sub("_", n)
    n = _ILLEGAL_PATTERN.sub("_", n)
    n = re.sub(r"\s+", " ", n).strip()
    n = n[:_MAX_SHEET]
    n = n.lstrip("_")
    if not n:
        n = "Sheet"
    return n


def dedupe(base: str, existing: set[str], max_length: int | None = _MAX_SHEET) -> str:
    if max_length is not None:
        base = base[:max_length]

    if base not in existing:
        existing.add(base)
        return base

    i = 2
    while True:
        suffix = f"_{i}"
        if max_length is None:
            candidate = f"{base}{suffix}"
        else:
            trim = max_length - len(suffix)
            core = base[:trim] if trim > 0 else ""
            candidate = f"{core}{suffix}"[-max_length:]
        if candidate not in existing and (max_length is None or len(candidate) <= max_length):
            existing.add(candidate)
            return candidate
        i += 1
