from collections.abc import Mapping
from contextlib import AbstractContextManager
from typing import Protocol

import pandas as pd


class SheetStream(Protocol):
    def append(self, df: pd.DataFrame) -> None: ...


class WorkbookWriter(Protocol):
    def write_single_sheet(self, df: pd.DataFrame, out_path: str, sheet_name: str = "Data") -> None: ...
    def write_multi_sheets(self, mapping: Mapping[str, pd.DataFrame], out_path: str) -> None: ...
    def stream_single_sheet(self, out_path: str, sheet_name: str = "Data") -> AbstractContextManager[SheetStream]: ...


class TableWriter(Protocol):
    def write_dataframe(
        self,
        df: pd.DataFrame,
        table: str,
        *,
        mode: str,
        options: Mapping[str, object] | None = None,
        uri: str,
    ) -> None: ...


class CloudObjectWriter(Protocol):
    def stream_object(
        self,
        key: str,
        format: str,
    ) -> AbstractContextManager[SheetStream]: ...
