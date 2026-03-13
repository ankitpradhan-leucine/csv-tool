"""Shared pytest fixtures."""

import pytest
from pathlib import Path


@pytest.fixture
def fixtures_dir() -> Path:
    """Return path to test fixtures directory."""
    return Path(__file__).parent / "fixtures"


@pytest.fixture
def sample_workbook_path(fixtures_dir: Path) -> Path:
    """Return path to sample test workbook."""
    return fixtures_dir / "sample_workbook.xlsx"
