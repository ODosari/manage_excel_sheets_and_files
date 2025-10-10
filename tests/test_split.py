from collections.abc import Mapping
from pathlib import Path

import pandas as pd
import pytest

from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.models import DatabaseDestination, SplitPlan
from excelmgr.core.split import split


class DummyReader:
    def read_sheet(self, path: str, sheet: str | int, password: str | None = None) -> pd.DataFrame:
        return pd.DataFrame({
            "Category": ["A/B", "A:B"],
            "Value": [1, 2],
        })


class DummyWriter:
    def __init__(self) -> None:
        self.sheet_calls: list[str] = []

    def write_single_sheet(self, df: pd.DataFrame, out_path: str, sheet_name: str = "Data") -> None:  # pragma: no cover - unused
        self.sheet_calls.append(sheet_name)

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
    assert sorted(Path(p).name for p in result["outputs"]) == files


def test_split_supports_index_spec_and_progress(tmp_path: Path) -> None:
    class IndexReader:
        def read_sheet(self, path: str, sheet: str | int, password: str | None = None) -> pd.DataFrame:
            return pd.DataFrame({
                "Category": [1, 2],
                "Value": [10, 20],
            })

    events: list[tuple[str, dict[str, object]]] = []

    def _hook(event: str, payload: dict[str, object]) -> None:
        events.append((event, payload))

    plan = SplitPlan(
        input_file="ignored.xlsx",
        by_column="index:0",
        to="files",
        output_dir=str(tmp_path),
        output_format="csv",
        dry_run=True,
    )

    result = split(plan, IndexReader(), DummyWriter(), progress_hooks=[_hook])

    assert [event for event, _ in events] == [
        "split_start",
        "split_partition",
        "split_partition",
        "split_complete",
    ]
    assert events[1][1]["rows"] == 1
    assert events[-1][1]["partitions"] == 2
    assert result["count"] == 2
    assert len(result["outputs"]) == 2


def test_split_respects_sheet_name_for_workbooks(tmp_path: Path) -> None:
    plan = SplitPlan(
        input_file="ignored.xlsx",
        by_column="Category",
        to="files",
        output_dir=str(tmp_path),
        output_format="xlsx",
        output_sheet_name="Report",
    )

    writer = DummyWriter()
    split(plan, DummyReader(), writer)

    assert writer.sheet_calls == ["Report", "Report"]


def test_split_custom_output_filename_for_sheets(tmp_path: Path) -> None:
    plan = SplitPlan(
        input_file="ignored.xlsx",
        by_column="Category",
        to="sheets",
        output_dir=str(tmp_path),
        output_filename="custom.xlsx",
        output_format="xlsx",
    )

    class SheetWriter:
        def __init__(self) -> None:
            self.received: tuple[dict[str, pd.DataFrame], str] | None = None

        def write_single_sheet(self, df: pd.DataFrame, out_path: str, sheet_name: str = "Data") -> None:  # pragma: no cover - unused
            raise AssertionError

        def write_multi_sheets(self, mapping, out_path: str) -> None:
            self.received = (dict(mapping), out_path)

        def stream_single_sheet(self, out_path: str, sheet_name: str = "Data"):  # pragma: no cover - unused
            raise AssertionError

    writer = SheetWriter()
    result = split(plan, DummyReader(), writer)

    assert result["out"].endswith("custom.xlsx")
    assert writer.received is not None
    assert writer.received[1].endswith("custom.xlsx")


def test_split_rejects_output_filename_for_files(tmp_path: Path) -> None:
    plan = SplitPlan(
        input_file="ignored.xlsx",
        by_column="Category",
        to="files",
        output_dir=str(tmp_path),
        output_filename="not-allowed.xlsx",
        output_format="xlsx",
    )

    with pytest.raises(ExcelMgrError):
        split(plan, DummyReader(), DummyWriter())


def test_split_writes_to_database_destination(tmp_path: Path) -> None:
    options = {"chunksize": 1000}
    destination = DatabaseDestination(
        uri=str(tmp_path / "parts.sqlite"),
        table="parts",
        mode="replace",
        options=options,
    )
    plan = SplitPlan(
        input_file="ignored.xlsx",
        by_column="Category",
        to="files",
        output_dir=str(tmp_path),
        output_format="csv",
        destination=destination,
    )

    class RecordingWriter:
        def __init__(self) -> None:
            self.calls: list[dict[str, object]] = []

        def write_dataframe(
            self,
            df: pd.DataFrame,
            table: str,
            *,
            mode: str,
            options: Mapping[str, object] | None = None,
            uri: str,
        ) -> None:
            self.calls.append(
                {
                    "df": df.copy(),
                    "table": table,
                    "mode": mode,
                    "options": dict(options or {}),
                    "uri": uri,
                }
            )

    writer = RecordingWriter()
    result = split(
        plan,
        DummyReader(),
        DummyWriter(),
        database_writer=writer,  # type: ignore[arg-type]
    )

    assert [call["mode"] for call in writer.calls] == ["replace", "append"]
    assert all(call["table"] == "parts" for call in writer.calls)
    assert all(call["uri"].endswith("parts.sqlite") for call in writer.calls)
    assert writer.calls[0]["options"] == options
    assert result["destination"] == {
        "kind": "database",
        "uri": destination.uri,
        "table": destination.table,
    }


def test_split_database_destination_requires_writer(tmp_path: Path) -> None:
    destination = DatabaseDestination(
        uri=str(tmp_path / "parts.sqlite"),
        table="parts",
    )
    plan = SplitPlan(
        input_file="ignored.xlsx",
        by_column="Category",
        to="files",
        output_dir=str(tmp_path),
        output_format="csv",
        destination=destination,
    )

    with pytest.raises(ExcelMgrError):
        split(plan, DummyReader(), DummyWriter())


def test_split_database_destination_requires_files_mode(tmp_path: Path) -> None:
    destination = DatabaseDestination(
        uri=str(tmp_path / "parts.sqlite"),
        table="parts",
    )
    plan = SplitPlan(
        input_file="ignored.xlsx",
        by_column="Category",
        to="sheets",
        output_dir=str(tmp_path),
        output_format="xlsx",
        destination=destination,
    )

    with pytest.raises(ExcelMgrError):
        split(plan, DummyReader(), DummyWriter())
