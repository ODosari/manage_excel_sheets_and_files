from contextlib import contextmanager
from pathlib import Path

import pandas as pd

from excelmgr.core.combine import combine
from excelmgr.core.models import CombinePlan


class StubReader:
    def __init__(self) -> None:
        self.sheet_calls: list[tuple[str, str | None]] = []
        self.read_calls: list[tuple[str, str | int, str | None]] = []

    def sheet_names(self, path: str, password: str | None = None) -> list[str]:
        self.sheet_calls.append((path, password))
        return ["Sheet1"]

    def read_sheet(self, path: str, sheet: str | int, password: str | None = None) -> pd.DataFrame:
        self.read_calls.append((path, sheet, password))
        return pd.DataFrame({"path": [path], "password": [password]})

    def read_workbook(self, path: str, password: str | None = None):  # pragma: no cover - unused here
        return {}

    def iter_files(self, root: str, glob: str | None, recursive: bool):  # pragma: no cover - unused here
        yield from []


class StubWriter:
    def __init__(self) -> None:
        self.appended: list[pd.DataFrame] = []
        self.multi_written: dict[str, pd.DataFrame] | None = None
        self.finalized = False
        self.multi_events: list[tuple[str, pd.DataFrame]] = []
        self.single_sheet_title: str | None = None

    def write_single_sheet(self, df: pd.DataFrame, out_path: str, sheet_name: str = "Data") -> None:  # pragma: no cover - not used
        raise AssertionError("write_single_sheet should not be called for streaming test")

    def write_multi_sheets(self, mapping, out_path: str) -> None:  # pragma: no cover - not used
        self.multi_written = mapping

    @contextmanager
    def stream_single_sheet(self, out_path: str, sheet_name: str = "Data"):
        self.single_sheet_title = sheet_name
        class _Recorder:
            def __init__(self, owner: "StubWriter") -> None:
                self._owner = owner

            def append(self, df: pd.DataFrame) -> None:
                self._owner.appended.append(df.copy())

            def finalize(self) -> None:
                self._owner.finalized = True

        recorder = _Recorder(self)
        try:
            yield recorder
        finally:
            recorder.finalize()

    @contextmanager
    def stream_multi_sheets(self, out_path: str):
        writer = self

        class _Recorder:
            def write_sheet(self, name: str, df: pd.DataFrame) -> None:
                writer.multi_events.append((name, df.copy()))

        yield _Recorder()


def test_combine_uses_password_map_and_streaming(tmp_path: Path) -> None:
    file_a = tmp_path / "a.xlsx"
    file_b = tmp_path / "b.xlsx"
    file_a.touch()
    file_b.touch()

    reader = StubReader()
    writer = StubWriter()

    plan = CombinePlan(
        inputs=[str(file_a), str(file_b)],
        include_sheets="all",
        output_path=str(tmp_path / "out.xlsx"),
        password="default",
        password_map={file_b.name: "special"},
        output_format="xlsx",
    )

    result = combine(plan, reader, writer)

    assert result["rows"] == 2
    assert writer.finalized is True
    assert [call[2] for call in reader.read_calls] == ["default", "special"]
    assert len(writer.appended) == 2
    assert all("password" in df.columns for df in writer.appended)


def test_combine_respects_sheet_name(tmp_path: Path) -> None:
    file_path = tmp_path / "single.xlsx"
    file_path.touch()

    reader = StubReader()
    writer = StubWriter()
    plan = CombinePlan(
        inputs=[str(file_path)],
        include_sheets="all",
        output_path=str(tmp_path / "out.xlsx"),
        output_format="xlsx",
        sheet_name="Summary",
    )

    combine(plan, reader, writer)

    assert writer.single_sheet_title == "Summary"


def test_combine_multi_sheet_streams(tmp_path: Path) -> None:
    file_path = tmp_path / "source.xlsx"
    file_path.touch()

    reader = StubReader()
    writer = StubWriter()
    plan = CombinePlan(
        inputs=[str(file_path)],
        include_sheets="all",
        mode="multi_sheets",
        output_path=str(tmp_path / "multi.xlsx"),
    )

    result = combine(plan, reader, writer)

    assert result["mode"] == "multi_sheets"
    assert writer.multi_events and writer.multi_events[0][0] == "Sheet1"
