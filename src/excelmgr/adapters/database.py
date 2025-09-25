from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
import sqlite3

import pandas as pd

from excelmgr.core.errors import ExcelMgrError


class SQLiteTableWriter:
    """Lightweight database writer backed by SQLite.

    The adapter satisfies the :class:`~excelmgr.ports.writers.TableWriter` protocol and
    is intentionally simple so that tests can persist data without external services.
    """

    def __init__(self, default_uri: str | None = None) -> None:
        self._default_uri = default_uri

    def write_dataframe(
        self,
        df: pd.DataFrame,
        table: str,
        *,
        mode: str,
        options: Mapping[str, object] | None = None,
        uri: str,
    ) -> None:
        target = uri or self._default_uri
        if not target:
            raise ExcelMgrError("SQLiteTableWriter requires a database 'uri' to be provided.")

        if mode not in {"replace", "append"}:
            raise ExcelMgrError(f"Unsupported database write mode: {mode}")

        path = Path(target).expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        kwargs = dict(options or {})
        with sqlite3.connect(path) as conn:
            df.to_sql(table, conn, if_exists=mode, index=False, **kwargs)
