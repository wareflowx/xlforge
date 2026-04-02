import pytest
from typer.testing import CliRunner

from xlforge import app


@pytest.fixture(scope="session")
def runner():
    """Pre-configured CliRunner instance."""
    return CliRunner()
