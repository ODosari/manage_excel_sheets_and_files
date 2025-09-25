import re

_MAX_SHEET = 31
_ILLEGAL = r'[:\\/?*\[\]]'

def sanitize_sheet_name(name: str) -> str:
    n = str(name).strip()
    if not n:
        n = "Empty"
    n = re.sub(_ILLEGAL, "_", n)
    n = re.sub(r"\s+", " ", n).strip()
    n = n.replace("/", "_").replace("\\", "_")
    n = n[:_MAX_SHEET]
    n = n.lstrip("_")
    if not n:
        n = "Sheet"
    return n

def dedupe(base: str, existing: set[str]) -> str:
    base = base[:_MAX_SHEET]
    if base not in existing:
        existing.add(base)
        return base

    i = 2
    while True:
        suffix = f"_{i}"
        trim = _MAX_SHEET - len(suffix)
        core = base[:trim] if trim > 0 else ""
        name = f"{core}{suffix}"[-_MAX_SHEET:]
        if name not in existing and len(name) <= _MAX_SHEET:
            existing.add(name)
            return name
        i += 1
