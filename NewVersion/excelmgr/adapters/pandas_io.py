from typing import Iterator, Mapping
import pandas as pd
from pathlib import Path
from excelmgr.adapters.local_storage import iter_files as _iter_files
from excelmgr.adapters.xls_protection import unlock_to_stream
from excelmgr.adapters.atomic import atomic_write
from excelmgr.config.settings import settings
from excelmgr.core.errors import MacroLossWarning, SheetNotFound
import warnings

class PandasReader:
    def __init__(self, engine: str = "openpyxl") -> None:
        self.engine = engine

    def sheet_names(self, path: str, password: str | None = None) -> list[str]:
        handle = path
        if password:
            handle = unlock_to_stream(path, password)
        with pd.ExcelFile(handle, engine=self.engine) as xf:
            return list(xf.sheet_names)

    def read_sheet(self, path: str, sheet: str | int, password: str | None = None) -> pd.DataFrame:
        handle = path
        if password:
            handle = unlock_to_stream(path, password)
        try:
            return pd.read_excel(handle, sheet_name=sheet, engine=self.engine)
        except ValueError as e:
            raise SheetNotFound(str(e))

    def iter_files(self, root: str, glob: str | None, recursive: bool) -> Iterator[str]:
        yield from _iter_files(root, glob or settings.glob, recursive)

class PandasWriter:
    def __init__(self, engine: str = "openpyxl") -> None:
        self.engine = engine

    def _macro_policy(self, out_path: str):
        if Path(out_path).suffix.lower() == ".xlsm":
            if settings.macro_policy == "warn":
                warnings.warn("Writing .xlsm will drop macros via openpyxl/pandas.", MacroLossWarning)
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
