import pytest
from typer.testing import CliRunner


@pytest.fixture(scope="session")
def runner():
    """Pre-configured CliRunner instance."""
    return CliRunner()
