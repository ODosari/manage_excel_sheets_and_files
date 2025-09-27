from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
from typer.testing import CliRunner

from excelmgr.cli.interactive import _compose_filename, pick_sheets
from excelmgr.cli.main import app
from excelmgr.core.password_maps import load_password_map
from excelmgr.core.sinks import csv_sink
from excelmgr.util.text import read_text, write_text


def test_filename_sanitize_unicode(tmp_path: Path) -> None:
    name = _compose_filename(
        base="ميزان/تجريبي",
        src=None,
        sheet="بيانات العملاء",
        extra="年報",
        timestamp="20250101_010101",
        suffix=".xlsx",
    )
    assert name == "ميزان_تجريبي__بيانات العملاء__年報__20250101_010101.xlsx"

    src = tmp_path / "اسم.xlsx"
    src.write_text("dummy", encoding="utf-8")
    deduped = _compose_filename(
        base=None,
        src=src,
        sheet="اسم",
        extra="اسم",
        timestamp="20250101_010101",
        suffix=".xlsx",
    )
    assert deduped.startswith("اسم__20250101_010101")


def test_csv_bom_handling(tmp_path: Path) -> None:
    bom_csv = tmp_path / "pw.csv"
    write_text(bom_csv, "path,password\nÜnicøde.xlsx,密码\n", add_bom=True)
    mapping = load_password_map(str(bom_csv))
    assert mapping == {"Ünicøde.xlsx": "密码"}

    df = pd.DataFrame({"Müşteri Adı": ["Ada"]})
    out_plain = tmp_path / "plain.csv"
    with csv_sink(str(out_plain)) as sink:
        sink.append(df)
    raw_plain = out_plain.read_bytes()
    assert not raw_plain.startswith("\ufeff".encode("utf-8"))
    assert read_text(out_plain).splitlines()[0] == "Müşteri Adı"

    out_bom = tmp_path / "bom.csv"
    with csv_sink(str(out_bom), add_bom=True) as sink:
        sink.append(df)
    raw_bom = out_bom.read_bytes()
    assert raw_bom.startswith("\ufeff".encode("utf-8"))


def test_sheet_picker_unicode(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    names = ["Résumé ✅", "بيانات العملاء"]
    monkeypatch.setattr("excelmgr.cli.interactive._list_sheet_names", lambda path, password: names)

    responses = iter(["1"])
    monkeypatch.setattr(
        "excelmgr.cli.interactive.typer.prompt",
        lambda message, **kwargs: next(responses),
    )
    selected = pick_sheets("dummy.xlsx", password=None, allow_multi=False, allow_all=False)
    assert selected == [names[0]]
    out = capsys.readouterr().out
    assert "1) Résumé ✅" in out

    responses = iter(["ré"])
    monkeypatch.setattr(
        "excelmgr.cli.interactive.typer.prompt",
        lambda message, **kwargs: next(responses),
    )
    selected_prefix = pick_sheets("dummy.xlsx", password=None, allow_multi=False, allow_all=False)
    assert selected_prefix == [names[0]]


def test_welcome_once() -> None:
    runner = CliRunner()
    result = runner.invoke(app, [], input="0\n8\n")
    assert result.exit_code == 0
    welcome_line = "Welcome to Excel Manager — a guided CLI for combining, splitting, previewing, and fixing Excel files."
    assert result.output.count(welcome_line) == 1
    assert result.output.count("Select an action") >= 2
