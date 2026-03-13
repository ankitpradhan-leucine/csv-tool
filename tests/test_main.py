"""Tests for CLI interface."""

import pytest
from pathlib import Path
from click.testing import CliRunner
from unittest.mock import patch, AsyncMock, MagicMock

from csvtool.main import cli, run


class TestCLI:
    """Tests for CLI commands."""

    @pytest.fixture
    def runner(self) -> CliRunner:
        """Create CLI test runner."""
        return CliRunner()

    def test_cli_help(self, runner: CliRunner):
        """Test CLI shows help."""
        result = runner.invoke(cli, ["--help"])
        assert result.exit_code == 0
        assert "CSV Automation Tool" in result.output

    def test_cli_version(self, runner: CliRunner):
        """Test CLI shows version."""
        result = runner.invoke(cli, ["--version"])
        assert result.exit_code == 0
        assert "0.1.0" in result.output

    def test_run_requires_url(self, runner: CliRunner):
        """Test run command requires URL."""
        result = runner.invoke(cli, ["run", "--workbook", "test.xlsx"])
        assert result.exit_code != 0
        assert "url" in result.output.lower() or "required" in result.output.lower()

    def test_run_requires_workbook(self, runner: CliRunner):
        """Test run command requires workbook."""
        result = runner.invoke(cli, ["run", "--url", "https://example.com"])
        assert result.exit_code != 0
        assert "workbook" in result.output.lower() or "required" in result.output.lower()

    def test_run_validates_workbook_exists(self, runner: CliRunner, tmp_path: Path):
        """Test run command validates workbook exists."""
        result = runner.invoke(cli, [
            "run",
            "--url", "https://example.com",
            "--workbook", "/nonexistent/workbook.xlsx"
        ])
        assert result.exit_code != 0
        assert "not found" in result.output.lower() or "does not exist" in result.output.lower()

    @patch("csvtool.main.Executor")
    @patch("csvtool.main.EvidenceGenerator")
    def test_run_executes_successfully(
        self,
        mock_evidence_gen,
        mock_executor,
        runner: CliRunner,
        tmp_path: Path
    ):
        """Test run command executes tests."""
        # Create temp workbook
        workbook = tmp_path / "test.xlsx"
        workbook.touch()

        # Mock executor
        mock_summary = MagicMock()
        mock_summary.total_tests = 5
        mock_summary.passed_tests = 5
        mock_summary.failed_tests = 0
        mock_summary.success_rate = 100.0

        mock_executor_instance = MagicMock()
        mock_executor_instance.run = AsyncMock(return_value=mock_summary)
        mock_executor.return_value = mock_executor_instance

        # Mock evidence generator
        mock_evidence_gen.return_value.generate.return_value = tmp_path / "evidence.docx"

        result = runner.invoke(cli, [
            "run",
            "--url", "https://example.com",
            "--workbook", str(workbook)
        ])

        assert result.exit_code == 0
        mock_executor_instance.run.assert_called_once()

    def test_run_with_output_dir(self, runner: CliRunner, tmp_path: Path):
        """Test run command accepts output directory."""
        workbook = tmp_path / "test.xlsx"
        workbook.touch()
        output_dir = tmp_path / "custom_output"

        with patch("csvtool.main.Executor") as mock_executor:
            mock_executor_instance = MagicMock()
            mock_executor_instance.run = AsyncMock(return_value=MagicMock(
                total_tests=1, passed_tests=1, failed_tests=0, success_rate=100.0
            ))
            mock_executor.return_value = mock_executor_instance

            with patch("csvtool.main.EvidenceGenerator") as mock_gen:
                mock_gen.return_value.generate.return_value = output_dir / "evidence.docx"

                result = runner.invoke(cli, [
                    "run",
                    "--url", "https://example.com",
                    "--workbook", str(workbook),
                    "--output", str(output_dir)
                ])

        assert result.exit_code == 0
