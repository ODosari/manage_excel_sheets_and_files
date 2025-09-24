from typing import Iterator
from pathlib import Path

def iter_files(root: str, glob: str | None = None, recursive: bool = False) -> Iterator[str]:
    p = Path(root)
    if p.is_file():
        yield str(p)
        return
    patterns = [g.strip() for g in (glob or "*.xlsx").split(",")]
    paths = p.rglob("*") if recursive else p.glob("*")
    for fp in paths:
        if not fp.is_file():
            continue
        if fp.name.startswith("~$"):
            continue
        for pat in patterns:
            if fp.match(pat):
                yield str(fp)
                break
