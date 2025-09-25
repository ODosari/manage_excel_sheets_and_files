import os
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from typing import IO, Any


@contextmanager
def atomic_write(path: str, mode: str = "wb", tmp_dir: str | None = None) -> Iterator[tuple[IO[Any], str]]:
    dname = os.path.dirname(os.path.abspath(path)) or "."
    os.makedirs(dname, exist_ok=True)
    temp_root = tmp_dir or dname
    os.makedirs(temp_root, exist_ok=True)

    dest_device = os.stat(dname).st_dev
    temp_device = os.stat(temp_root).st_dev
    if dest_device != temp_device:
        raise OSError(
            "Temporary directory must reside on the same filesystem as the destination for atomic writes."
        )

    fd, tmp = tempfile.mkstemp(dir=temp_root, prefix=".tmp-", suffix=".partial")
    try:
        with os.fdopen(fd, mode) as f:
            yield f, tmp
        # Move into place atomically
        os.replace(tmp, path)
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
