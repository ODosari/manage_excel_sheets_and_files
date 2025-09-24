from typer.testing import CliRunner
from NewVersion.excelmgr.cli.main import app

def test_cli_version_and_help():
    r = CliRunner().invoke(app, ["version"])
    assert r.exit_code == 0
    r = CliRunner().invoke(app, ["--help"])
    assert r.exit_code == 0
