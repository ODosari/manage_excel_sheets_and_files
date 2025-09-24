from __future__ import annotations

from contextlib import contextmanager
from typing import Iterator

import pandas as pd

from src.adapters.atomic import atomic_write
from src.config.settings import settings
from src.core.errors import ExcelMgrError


class _CsvSink:
    def __init__(self, handle) -> None:
        self._handle = handle
        self._header_written = False
        self._columns: list[str] | None = None

    def append(self, df: pd.DataFrame) -> None:
        if self._columns is None:
            self._columns = list(df.columns)
        if not self._header_written:
            header_frame = pd.DataFrame(columns=self._columns)
            header_frame.to_csv(self._handle, index=False)
            self._header_written = True
        if not df.empty:
            df.to_csv(self._handle, index=False, header=False)

    def finalize(self) -> None:
        if not self._header_written and self._columns is not None:
            header_frame = pd.DataFrame(columns=self._columns)
            header_frame.to_csv(self._handle, index=False)


class _ParquetSink:
    def __init__(self, tmp_path: str, pa_module, pq_module) -> None:
        self._tmp_path = tmp_path
        self._pa = pa_module
        self._pq = pq_module
        self._writer = None
        self._empty_columns: list[str] | None = None

    def append(self, df: pd.DataFrame) -> None:
        if df.empty:
            if self._writer is None and self._empty_columns is None:
                self._empty_columns = list(df.columns)
            return

        table = self._pa.Table.from_pandas(df, preserve_index=False)
        if self._writer is None:
            self._writer = self._pq.ParquetWriter(self._tmp_path, table.schema)
        self._writer.write_table(table)

    def finalize(self) -> None:
        if self._writer is not None:
            self._writer.close()
            return
        frame = pd.DataFrame(columns=self._empty_columns)
        frame.to_parquet(self._tmp_path, index=False)


@contextmanager
def csv_sink(out_path: str) -> Iterator[_CsvSink]:
    with atomic_write(out_path, "w", tmp_dir=settings.temp_dir) as (fh, _tmp):
        sink = _CsvSink(fh)
        try:
            yield sink
        finally:
            sink.finalize()


@contextmanager
def parquet_sink(out_path: str) -> Iterator[_ParquetSink]:
    try:
        import pyarrow as pa  # type: ignore[import-not-found]
        import pyarrow.parquet as pq  # type: ignore[import-not-found]
    except ImportError as exc:  # pragma: no cover - requires optional dependency
        raise ExcelMgrError("Writing parquet output requires the 'pyarrow' package to be installed.") from exc

    with atomic_write(out_path, "wb", tmp_dir=settings.temp_dir) as (fh, tmp):
        fh.close()
        sink = _ParquetSink(tmp, pa, pq)
        try:
            yield sink
        finally:
            sink.finalize()
