import json
from pathlib import Path

import pandas as pd
from typer.testing import CliRunner

from excelmgr.cli.main import app
from excelmgr.config.settings import settings

runner = CliRunner()


def _read_json(stdout: str) -> dict:
    lines = [line for line in stdout.splitlines() if line.strip()]
    return json.loads("\n".join(lines))


def test_cli_combine_creates_output_and_respects_temp_dir():
    with runner.isolated_filesystem():
        base = Path.cwd()
        data_dir = base / "inputs"
        data_dir.mkdir()
        df1 = pd.DataFrame({"Customer": ["A", "B"], "Amount": [10, 20]})
        df2 = pd.DataFrame({"Customer": ["C"], "Amount": [30]})
        df1.to_excel(data_dir / "north.xlsx", index=False)
        df2.to_excel(data_dir / "south.xlsx", index=False)

        original_temp = settings.temp_dir
        temp_dir = base / "tmp" / "excel"
        settings.temp_dir = str(temp_dir)
        try:
            result = runner.invoke(
                app,
                [
                    "combine",
                    str(data_dir / "north.xlsx"),
                    str(data_dir / "south.xlsx"),
                    "--out",
                    "output/reports/combined.xlsx",
                ],
                catch_exceptions=False,
            )
        finally:
            settings.temp_dir = original_temp

        assert result.exit_code == 0, result.stdout
        payload = _read_json(result.stdout)
        out_file = base / "output" / "reports" / "combined.xlsx"
        assert out_file.exists()
        combined = pd.read_excel(out_file)
        assert len(combined) == 3
        assert payload["files"] == 2
        assert temp_dir.exists()


def test_cli_combine_to_csv_format():
    with runner.isolated_filesystem():
        df1 = pd.DataFrame({"Customer": ["A"], "Amount": [10]})
        df2 = pd.DataFrame({"Customer": ["B"], "Amount": [20]})
        df1.to_excel("north.xlsx", index=False, sheet_name="North")
        df2.to_excel("south.xlsx", index=False, sheet_name="South")

        result = runner.invoke(
            app,
            [
                "combine",
                "north.xlsx",
                "south.xlsx",
                "--out",
                "combined.csv",
                "--format",
                "csv",
            ],
            catch_exceptions=False,
        )

        assert result.exit_code == 0, result.stdout
        payload = _read_json(result.stdout)
        out_file = Path("combined.csv")
        assert out_file.exists()
        combined = pd.read_csv(out_file)
        assert len(combined) == 2
        assert payload["format"] == "csv"


def test_cli_combine_dry_run_skips_output():
    with runner.isolated_filesystem():
        df1 = pd.DataFrame({"Customer": ["A"]})
        df2 = pd.DataFrame({"Customer": ["B"]})
        df1.to_excel("east.xlsx", index=False)
        df2.to_excel("west.xlsx", index=False)

        result = runner.invoke(
            app,
            [
                "combine",
                "east.xlsx",
                "west.xlsx",
                "--dry-run",
            ],
            catch_exceptions=False,
        )

        assert result.exit_code == 0, result.stdout
        assert not Path("combined.xlsx").exists()
        payload = _read_json(result.stdout)
        assert payload["dry_run"] is True


def test_cli_split_to_files_creates_directory_structure():
    with runner.isolated_filesystem():
        df = pd.DataFrame(
            {
                "Category": ["A/B", "A:B", "A/B"],
                "Value": [1, 2, 3],
            }
        )
        df.to_excel("source.xlsx", index=False)

        result = runner.invoke(
            app,
            [
                "split",
                "source.xlsx",
                "--sheet",
                "0",
                "--by",
                "Category",
                "--to",
                "files",
                "--out",
                "exports/items",
            ],
            catch_exceptions=False,
        )

        assert result.exit_code == 0, result.stdout
        payload = _read_json(result.stdout)
        a_file = Path("exports/items/A_B.xlsx")
        b_file = Path("exports/items/A_B_2.xlsx")
        assert a_file.exists() and b_file.exists()
        assert payload["count"] == 2
        assert sorted(Path(p).name for p in payload["outputs"]) == [
            "A_B.xlsx",
            "A_B_2.xlsx",
        ]
        assert pd.read_excel(a_file)["Category"].unique().tolist() == ["A/B"]


def test_cli_split_to_csv_format():
    with runner.isolated_filesystem():
        df = pd.DataFrame({"Category": ["A", "B"], "Value": [1, 2]})
        df.to_excel("source.xlsx", index=False)

        result = runner.invoke(
            app,
            [
                "split",
                "source.xlsx",
                "--by",
                "Category",
                "--to",
                "files",
                "--out",
                "exports",
                "--format",
                "csv",
            ],
            catch_exceptions=False,
        )

        assert result.exit_code == 0, result.stdout
        payload = _read_json(result.stdout)
        a_file = Path("exports/A.csv")
        b_file = Path("exports/B.csv")
        assert a_file.exists() and b_file.exists()
        assert payload["format"] == "csv"
        assert sorted(Path(p).name for p in payload["outputs"]) == [
            "A.csv",
            "B.csv",
        ]
        assert pd.read_csv(a_file)["Category"].tolist() == ["A"]


def test_cli_split_to_sheets_derives_output_name():
    with runner.isolated_filesystem():
        df = pd.DataFrame({"Group": ["G1", "G2"], "Value": [1, 2]})
        df.to_excel("mydata.xlsx", index=False)

        result = runner.invoke(
            app,
            [
                "split",
                "mydata.xlsx",
                "--by",
                "Group",
                "--to",
                "sheets",
                "--out",
                "exports",
            ],
            catch_exceptions=False,
        )

        assert result.exit_code == 0, result.stdout
        out_file = Path("exports/mydata_split.xlsx")
        assert out_file.exists()
        payload = _read_json(result.stdout)
        assert payload["dry_run"] is False
        assert payload["out"].endswith("mydata_split.xlsx")


def test_cli_split_honors_custom_output_filename_and_sheet_name():
    with runner.isolated_filesystem():
        df = pd.DataFrame({"Group": ["G1"], "Value": [1]})
        df.to_excel("data.xlsx", index=False)

        result = runner.invoke(
            app,
            [
                "split",
                "data.xlsx",
                "--by",
                "Group",
                "--to",
                "sheets",
                "--out",
                "exports",
                "--out-file",
                "custom.xlsx",
            ],
            catch_exceptions=False,
        )

        assert result.exit_code == 0, result.stdout
        payload = _read_json(result.stdout)
        assert payload["out"].endswith("custom.xlsx")
        assert Path("exports/custom.xlsx").exists()

        result_files = runner.invoke(
            app,
            [
                "split",
                "data.xlsx",
                "--by",
                "Group",
                "--to",
                "files",
                "--out",
                "exports",
                "--sheet-name",
                "Summary",
            ],
            catch_exceptions=False,
        )

        assert result_files.exit_code == 0, result_files.stdout
        payload_files = _read_json(result_files.stdout)
        first_output = Path(payload_files["outputs"][0])
        assert first_output.exists()
        default_df = pd.read_excel(first_output)
        assert default_df.columns.tolist()[0] == "Group"
        summary_df = pd.read_excel(first_output, sheet_name="Summary")
        assert summary_df.equals(default_df)


def test_cli_preview_reports_metadata():
    with runner.isolated_filesystem():
        with pd.ExcelWriter("report.xlsx") as writer:
            pd.DataFrame({"A": [1, 2], "B": [3, 4]}).to_excel(writer, sheet_name="Summary", index=False)
            pd.DataFrame({"X": ["x"]}).to_excel(writer, sheet_name="Notes", index=False)

        result = runner.invoke(app, ["preview", "report.xlsx"], catch_exceptions=False)

        assert result.exit_code == 0, result.stdout
        payload = _read_json(result.stdout)
        assert payload["path"].endswith("report.xlsx")
        assert {sheet["name"] for sheet in payload["sheets"]} == {"Summary", "Notes"}
        summary = next(sheet for sheet in payload["sheets"] if sheet["name"] == "Summary")
        assert summary["rows"] == 2
        assert summary["headers"] == ["A", "B"]


def test_cli_split_reports_missing_column_error():
    with runner.isolated_filesystem():
        df = pd.DataFrame({"Category": ["A", "B"], "Value": [1, 2]})
        df.to_excel("source.xlsx", index=False)

        result = runner.invoke(
            app,
            [
                "split",
                "source.xlsx",
                "--by",
                "Missing",
            ],
        )

        assert result.exit_code == 2
        message = result.stderr or result.stdout
        assert "Column 'Missing' was not found" in message


def test_cli_delete_cols_inplace_and_missing_error():
    with runner.isolated_filesystem():
        df = pd.DataFrame({"Keep": [1, 2], "DropMe": [3, 4]})
        df.to_excel("data.xlsx", index=False)

        ok = runner.invoke(
            app,
            [
                "delete-cols",
                "data.xlsx",
                "--targets",
                "DropMe",
                "--yes",
                "--inplace",
            ],
            catch_exceptions=False,
        )
        assert ok.exit_code == 0, ok.stdout
        cleaned = pd.read_excel("data.xlsx")
        assert "DropMe" not in cleaned.columns

        bad = runner.invoke(
            app,
            [
                "delete-cols",
                "data.xlsx",
                "--targets",
                "Missing",
                "--yes",
                "--on-missing",
                "error",
            ],
        )
        assert bad.exit_code == 2
        assert "Columns not found" in (bad.stderr or "")


def test_cli_delete_cols_confirmation_abort():
    with runner.isolated_filesystem():
        df = pd.DataFrame({"Keep": [1], "DropMe": [2]})
        df.to_excel("data.xlsx", index=False)

        result = runner.invoke(
            app,
            [
                "delete-cols",
                "data.xlsx",
                "--targets",
                "DropMe",
            ],
            input="n\n",
        )
        assert result.exit_code == 0
        assert not Path("data.cleaned.xlsx").exists()


def test_cli_delete_cols_defaults_to_single_sheet():
    with runner.isolated_filesystem():
        with pd.ExcelWriter("workbook.xlsx") as writer:
            pd.DataFrame({"Keep": [1, 2], "Drop": [3, 4]}).to_excel(
                writer, sheet_name="First", index=False
            )
            pd.DataFrame({"Keep": [5, 6], "Drop": [7, 8]}).to_excel(
                writer, sheet_name="Second", index=False
            )

        result = runner.invoke(
            app,
            [
                "delete-cols",
                "workbook.xlsx",
                "--targets",
                "Drop",
            ],
            input="y\n",
            catch_exceptions=False,
        )

        assert result.exit_code == 0, result.stdout
        out_file = Path("workbook.cleaned.xlsx")
        assert out_file.exists()
        first = pd.read_excel(out_file, sheet_name="First")
        second = pd.read_excel(out_file, sheet_name="Second")
        assert "Drop" not in first.columns
        assert "Drop" in second.columns


def test_cli_delete_cols_sheet_index_updates_target_only():
    with runner.isolated_filesystem():
        with pd.ExcelWriter("zones.xlsx") as writer:
            pd.DataFrame({"Keep": [1, 2], "Remove": [3, 4]}).to_excel(
                writer, sheet_name="North", index=False
            )
            pd.DataFrame({"Keep": [5, 6], "Remove": [7, 8]}).to_excel(
                writer, sheet_name="South", index=False
            )

        result = runner.invoke(
            app,
            [
                "delete-cols",
                "zones.xlsx",
                "--targets",
                "Remove",
                "--sheet",
                "1",
                "--yes",
            ],
            catch_exceptions=False,
        )

        assert result.exit_code == 0, result.stdout
        out_file = Path("zones.cleaned.xlsx")
        assert out_file.exists()
        data = pd.read_excel(out_file, sheet_name=None)
        assert set(data.keys()) == {"North", "South"}
        assert "Remove" in data["North"].columns
        assert "Remove" not in data["South"].columns


def test_cli_plan_executes_operations():
    with runner.isolated_filesystem():
        df1 = pd.DataFrame({"Region": ["North"], "Value": [10]})
        df2 = pd.DataFrame({"Region": ["South"], "Value": [20]})
        df1.to_excel("north.xlsx", index=False, sheet_name="North")
        df2.to_excel("south.xlsx", index=False, sheet_name="South")

        plan_text = """
operations:
  - type: combine
    name: combine_cloud
    options:
      inputs:
        - north.xlsx
        - south.xlsx
      output_format: csv
      output_path: combined_local.csv
      destination:
        kind: cloud
        root: cloud-bucket
        key: combined.csv
        format: csv
  - type: preview
    name: inspect
    options:
      path: north.xlsx
      limit: 1
"""
        Path("batch.yaml").write_text(plan_text, encoding="utf-8")

        result = runner.invoke(app, ["plan", "batch.yaml"], catch_exceptions=False)

        assert result.exit_code == 0, result.stdout
        payload = _read_json(result.stdout)
        operations = payload["operations"]
        assert [op["type"] for op in operations] == ["combine", "preview"]
        assert Path("combined_local.csv").exists()
        assert Path("cloud-bucket/combined.csv").exists()
        combine_result = operations[0]["result"]
        assert combine_result["destination"]["kind"] == "cloud"
        preview_result = operations[1]["result"]
        assert preview_result["sheets"][0]["name"] == "North"
