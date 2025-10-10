import warnings
from collections.abc import Iterator, Mapping
from contextlib import contextmanager
from pathlib import Path

import pandas as pd

from excelmgr.adapters.atomic import atomic_write
from excelmgr.adapters.local_storage import iter_files as _iter_files
from excelmgr.adapters.xls_protection import unlock_to_stream
from excelmgr.config.settings import settings
from excelmgr.core.errors import MacroLossWarning, SheetNotFound
from excelmgr.ports.writers import MultiSheetStream


class PandasReader:
    def __init__(self, engine: str = "openpyxl") -> None:
        self.engine = engine

    def sheet_names(self, path: str, password: str | None = None) -> list[str]:
        handle = path
        if password:
            handle = unlock_to_stream(path, password)
        with pd.ExcelFile(handle, engine=self.engine) as xf:
            return list(xf.sheet_names)

    def sheet_columns(self, path: str, sheet: str | int, password: str | None = None) -> list[object]:
        """Return the column labels for a sheet without loading all rows."""

        handle = path
        if password:
            handle = unlock_to_stream(path, password)
        with pd.ExcelFile(handle, engine=self.engine) as xf:
            frame = xf.parse(sheet_name=sheet, nrows=0)
        return list(frame.columns)

    def read_sheet(self, path: str, sheet: str | int, password: str | None = None) -> pd.DataFrame:
        handle = path
        if password:
            handle = unlock_to_stream(path, password)
        try:
            return pd.read_excel(handle, sheet_name=sheet, engine=self.engine)
        except ValueError as exc:
            raise SheetNotFound(str(exc)) from exc

    def read_workbook(self, path: str, password: str | None = None) -> Mapping[str, pd.DataFrame]:
        handle = path
        if password:
            handle = unlock_to_stream(path, password)
        with pd.ExcelFile(handle, engine=self.engine) as xf:
            return {name: xf.parse(sheet_name=name) for name in xf.sheet_names}

    def iter_files(self, root: str, glob: str | None, recursive: bool) -> Iterator[str]:
        yield from _iter_files(root, glob or settings.glob, recursive)

class PandasWriter:
    def __init__(self, engine: str = "openpyxl") -> None:
        self.engine = engine

    def _macro_policy(self, out_path: str):
        if Path(out_path).suffix.lower() == ".xlsm":
            if settings.macro_policy == "warn":
                warnings.warn(
                    "Writing .xlsm will drop macros via openpyxl/pandas.",
                    MacroLossWarning,
                    stacklevel=2,
                )
            elif settings.macro_policy == "forbid":
                raise MacroLossWarning("Refusing to write .xlsm: would drop macros.")
            # ignore => do nothing

    @staticmethod
    def _ensure_parent_dir(out_path: str) -> None:
        Path(out_path).expanduser().resolve().parent.mkdir(parents=True, exist_ok=True)

    def write_single_sheet(self, df: pd.DataFrame, out_path: str, sheet_name: str = "Data") -> None:
        self._macro_policy(out_path)
        self._ensure_parent_dir(out_path)
        with atomic_write(out_path, "wb", tmp_dir=settings.temp_dir) as (f, tmp):
            with pd.ExcelWriter(f, engine=self.engine) as w:
                df.to_excel(w, index=False, sheet_name=sheet_name)

    def write_multi_sheets(self, mapping: Mapping[str, pd.DataFrame], out_path: str) -> None:
        self._macro_policy(out_path)
        self._ensure_parent_dir(out_path)
        with atomic_write(out_path, "wb", tmp_dir=settings.temp_dir) as (f, tmp):
            with pd.ExcelWriter(f, engine=self.engine) as w:
                for name, df in mapping.items():
                    df.to_excel(w, index=False, sheet_name=name)

    @contextmanager
    def stream_single_sheet(self, out_path: str, sheet_name: str = "Data"):
        self._macro_policy(out_path)
        self._ensure_parent_dir(out_path)

        class _SheetAppender:
            def __init__(self, excel_writer: pd.ExcelWriter, target: str) -> None:
                self._writer = excel_writer
                self._sheet = target
                self._row = 0
                self._header_written = False

            def append(self, df: pd.DataFrame) -> None:
                header = not self._header_written
                to_write = df if not (header and df.empty) else df.head(0)
                if to_write.empty and not header:
                    return
                to_write.to_excel(
                    self._writer,
                    index=False,
                    sheet_name=self._sheet,
                    startrow=self._row,
                    header=header,
                )
                header_rows = 1 if header else 0
                self._row += header_rows + len(to_write)
                self._header_written = True

            def finalize(self) -> None:
                if not self._header_written:
                    pd.DataFrame().to_excel(
                        self._writer,
                        index=False,
                        sheet_name=self._sheet,
                    )

        with atomic_write(out_path, "wb", tmp_dir=settings.temp_dir) as (f, tmp):
            with pd.ExcelWriter(f, engine=self.engine) as w:
                appender = _SheetAppender(w, sheet_name)
                try:
                    yield appender
                finally:
                    appender.finalize()

    @contextmanager
    def stream_multi_sheets(self, out_path: str) -> Iterator[MultiSheetStream]:
        self._macro_policy(out_path)
        self._ensure_parent_dir(out_path)

        class _WorkbookAppender:
            def __init__(self, excel_writer: pd.ExcelWriter) -> None:
                self._writer = excel_writer
                self._row_positions: dict[str, int] = {}

            def append(self, sheet_name: str, df: pd.DataFrame) -> None:
                start = self._row_positions.get(sheet_name, 0)
                header = start == 0
                to_write = df if not (header and df.empty) else df.head(0)
                if to_write.empty and not header:
                    return
                to_write.to_excel(
                    self._writer,
                    index=False,
                    sheet_name=sheet_name,
                    startrow=start,
                    header=header,
                )
                header_rows = 1 if header else 0
                self._row_positions[sheet_name] = start + header_rows + len(to_write)
                if header and df.empty:
                    # Ensure the sheet materializes even when empty by writing a blank sheet
                    pd.DataFrame().to_excel(
                        self._writer,
                        index=False,
                        sheet_name=sheet_name,
                    )

        with atomic_write(out_path, "wb", tmp_dir=settings.temp_dir) as (f, tmp):
            with pd.ExcelWriter(f, engine=self.engine) as w:
                appender = _WorkbookAppender(w)
                yield appender
