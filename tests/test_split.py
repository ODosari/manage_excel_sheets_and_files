from pathlib import Path

import pandas as pd

from excelmgr.core.models import SplitPlan
from excelmgr.core.split import split


class DummyReader:
    def read_sheet(self, path: str, sheet: str | int, password: str | None = None) -> pd.DataFrame:
        return pd.DataFrame({
            "Category": ["A/B", "A:B"],
            "Value": [1, 2],
        })


class DummyWriter:
    def write_single_sheet(self, df: pd.DataFrame, out_path: str, sheet_name: str = "Data") -> None:  # pragma: no cover - unused
        raise AssertionError("write_single_sheet should not be called for CSV split")

    def write_multi_sheets(self, mapping, out_path: str) -> None:  # pragma: no cover - unused
        raise AssertionError("write_multi_sheets should not be called for CSV split")

    def stream_single_sheet(self, out_path: str, sheet_name: str = "Data"):  # pragma: no cover - unused
        raise AssertionError("stream_single_sheet should not be called for CSV split")


def test_split_dedupes_file_names(tmp_path: Path) -> None:
    plan = SplitPlan(
        input_file="ignored.xlsx",
        by_column="Category",
        to="files",
        output_dir=str(tmp_path),
        output_format="csv",
        dry_run=False,
    )

    result = split(plan, DummyReader(), DummyWriter())

    files = sorted(p.name for p in tmp_path.iterdir())
    assert files == ["A_B.csv", "A_B_2.csv"]
    assert result["count"] == 2
