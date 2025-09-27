from __future__ import annotations

from typer.testing import CliRunner

from excelmgr.cli.main import app


def test_default_launches_interactive_menu():
    runner = CliRunner()
    result = runner.invoke(app, [], input="8\n")
    assert result.exit_code == 0
    assert "Main menu" in result.output


def test_interactive_combine_flow(monkeypatch):
    captured = {}

    def fake_reader():
        return "reader"

    def fake_writer():
        return "writer"

    def fake_combine(plan, reader, writer, progress_hooks):  # pragma: no cover - invoked via interactive menu
        captured["plan"] = plan
        return {"output_path": plan.output_path}

    monkeypatch.setattr("excelmgr.cli.interactive.PandasReader", fake_reader)
    monkeypatch.setattr("excelmgr.cli.interactive.PandasWriter", fake_writer)
    monkeypatch.setattr("excelmgr.cli.interactive.combine_command", fake_combine)

    runner = CliRunner()
    input_steps = [
        "1",  # main menu -> combine
        "file.xlsx",  # inputs
        "1",  # mode one-sheet
        "",  # glob
        "n",  # recursive
        "",  # sheets default all
        "",  # output path default
        "",  # sheet name default
        "n",  # add source column
        "1",  # format xlsx
        "y",  # dry run
        "5",  # password menu -> done
        "2",  # try again/back menu -> back
        "8",  # exit
    ]
    result = runner.invoke(app, [], input="\n".join(input_steps) + "\n")
    assert result.exit_code == 0
    assert "Combine workflow" in result.output
    assert captured["plan"].inputs == ["file.xlsx"]
    assert captured["plan"].mode == "one_sheet"
    assert captured["plan"].dry_run is True
