from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.sinks import csv_sink, parquet_sink
from excelmgr.ports.writers import SheetStream


class LocalCloudObjectWriter:
    """Trivial object-storage writer backed by the local filesystem.

    The adapter mimics uploading objects to a bucket by writing the output files into
    a directory.  It implements the :class:`~excelmgr.ports.writers.CloudObjectWriter`
    protocol so the core logic can remain unaware of the actual storage backend.
    """

    def __init__(self, root: str) -> None:
        self._root = Path(root).expanduser().resolve()
        self._root.mkdir(parents=True, exist_ok=True)

    @contextmanager
    def stream_object(self, key: str, format: str) -> SheetStream:
        fmt = format.lower()
        path = (self._root / key).expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)

        if fmt == "csv":
            with csv_sink(str(path)) as sink:
                yield sink
            return
        if fmt == "parquet":
            with parquet_sink(str(path)) as sink:
                yield sink
            return
        if fmt == "xlsx":
            from excelmgr.adapters.pandas_io import PandasWriter

            writer = PandasWriter()
            with writer.stream_single_sheet(str(path), sheet_name="Data") as sink:
                yield sink
            return

        raise ExcelMgrError(f"Unsupported cloud object format: {format}")
