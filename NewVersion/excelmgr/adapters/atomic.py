import os, tempfile, shutil
from contextlib import contextmanager

@contextmanager
def atomic_write(path: str, mode: str = "wb"):
    dname = os.path.dirname(os.path.abspath(path)) or "."
    fd, tmp = tempfile.mkstemp(dir=dname, prefix=".tmp-", suffix=".partial")
    try:
        with os.fdopen(fd, mode) as f:
            yield f, tmp
        # Move into place atomically
        shutil.move(tmp, path)
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
