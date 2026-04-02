from typer.testing import CliRunner
from xlforge import app

runner = CliRunner()


def test_ping():
    result = runner.invoke(app, ["ping"])
    assert result.exit_code == 0
    assert "pong" in result.output


def test_version():
    result = runner.invoke(app, ["version"])
    assert result.exit_code == 0
    assert "xlforge 0.1.0" in result.output


def test_no_args_shows_help():
    result = runner.invoke(app)
    # Without invoke_without_command, no args shows help/missing command error
    assert result.exit_code in [0, 2]
