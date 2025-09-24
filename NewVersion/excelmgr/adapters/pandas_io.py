from typing import Iterator, Mapping
import pandas as pd
from pathlib import Path
from NewVersion.excelmgr.adapters.local_storage import iter_files as _iter_files
from NewVersion.excelmgr.adapters.xls_protection import unlock_to_stream
from NewVersion.excelmgr.adapters.atomic import atomic_write
from NewVersion.excelmgr.config.settings import settings
from NewVersion.excelmgr.core.errors import MacroLossWarning, SheetNotFound
import warnings

class PandasReader:
    @staticmethod
    def sheet_names(self, path: str, password: str | None = None) -> list[str]:
        handle = path
        if password:
            handle = unlock_to_stream(path, password)
        with pd.ExcelFile(handle, engine="openpyxl") as xf:
            return list(xf.sheet_names)

    @staticmethod
    def read_sheet(self, path: str, sheet: str | int, password: str | None = None) -> pd.DataFrame:
        handle = path
        if password:
            handle = unlock_to_stream(path, password)
        try:
            return pd.read_excel(handle, sheet_name=sheet, engine="openpyxl")
        except ValueError as e:
            raise SheetNotFound(str(e))

    @staticmethod
    def iter_files(self, root: str, glob: str | None, recursive: bool) -> Iterator[str]:
        yield from _iter_files(root, glob or settings.glob, recursive)

class PandasWriter:
    @staticmethod
    def _macro_policy(self, out_path: str):
        if Path(out_path).suffix.lower() == ".xlsm":
            if settings.macro_policy == "warn":
                warnings.warn("Writing .xlsm will drop macros via openpyxl/pandas.", MacroLossWarning)
            elif settings.macro_policy == "forbid":
                raise MacroLossWarning("Refusing to write .xlsm: would drop macros.")
            # ignore => do nothing

    def write_single_sheet(self, df: pd.DataFrame, out_path: str, sheet_name: str = "Data") -> None:
        self._macro_policy(out_path)
        with atomic_write(out_path, "wb") as (f, tmp):
            with pd.ExcelWriter(f, engine="openpyxl") as w:
                df.to_excel(w, index=False, sheet_name=sheet_name)

    def write_multi_sheets(self, mapping: Mapping[str, pd.DataFrame], out_path: str) -> None:
        self._macro_policy(out_path)
        with atomic_write(out_path, "wb") as (f, tmp):
            with pd.ExcelWriter(f, engine="openpyxl") as w:
                for name, df in mapping.items():
                    df.to_excel(w, index=False, sheet_name=name)
