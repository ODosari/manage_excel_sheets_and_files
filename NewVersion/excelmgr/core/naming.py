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
    if base not in existing:
        existing.add(base)
        return base
    i = 2
    while True:
        name = f"{base}_{i}"
        if name not in existing and len(name) <= _MAX_SHEET:
            existing.add(name)
            return name
        i += 1
