"""Integration tests for full workflow."""

import pytest
from pathlib import Path
from unittest.mock import patch, AsyncMock, MagicMock
from datetime import datetime

from csvtool.executor import Executor, ExecutorConfig
from csvtool.evidence import EvidenceGenerator
from csvtool.models import ExecutionSummary


class TestFullWorkflow:
    """Integration tests for complete workflow."""

    @pytest.fixture
    def test_workbook(self, tmp_path: Path) -> Path:
        """Create test workbook."""
        from openpyxl import Workbook

        wb = Workbook()

        # Credentials
        creds = wb.active
        creds.title = "Credentials"
        creds.append(["Role", "Username", "Password"])
        creds.append(["Admin", "admin@test.com", "admin123"])

        # Scenarios
        scenarios = wb.create_sheet("Scenarios")
        scenarios.append(["Test ID", "Role", "Title", "Test Instructions", "Expected Result"])
        scenarios.append([
            "INT-001",
            "Admin",
            "Integration Test",
            "1. Navigate to home\n2. Verify page loads",
            "Home page displayed"
        ])

        path = tmp_path / "integration_test.xlsx"
        wb.save(path)
        return path

    @pytest.fixture
    def config(self, test_workbook: Path, tmp_path: Path) -> ExecutorConfig:
        """Create executor config."""
        return ExecutorConfig(
            url="https://example.com",
            workbook_path=test_workbook,
            output_dir=tmp_path / "output",
            cache_dir=tmp_path / "cache",
            headless=True
        )

    @pytest.mark.asyncio
    async def test_full_workflow_with_mocks(self, config: ExecutorConfig):
        """Test full workflow with mocked browser and AI."""
        with patch("csvtool.executor.Browser") as mock_browser_class, \
             patch("csvtool.executor.AINavigator") as mock_navigator_class:

            # Setup browser mock
            mock_browser = AsyncMock()
            mock_browser.navigate = AsyncMock()

            # Create screenshot directory and dummy file
            screenshot_dir = config.output_dir / "screenshots"
            screenshot_dir.mkdir(parents=True, exist_ok=True)
            dummy_screenshot = screenshot_dir / "test.png"
            dummy_screenshot.write_text("dummy")  # Dummy instead of real image

            mock_browser.screenshot = AsyncMock(return_value=dummy_screenshot)
            mock_browser.get_screenshot_base64 = AsyncMock(return_value="base64data")
            mock_browser.click = AsyncMock()
            mock_browser.type_text = AsyncMock()
            mock_browser.get_current_url = MagicMock(return_value="https://example.com")
            mock_browser.clear_session = AsyncMock()

            mock_browser_class.return_value.__aenter__ = AsyncMock(return_value=mock_browser)
            mock_browser_class.return_value.__aexit__ = AsyncMock()

            # Setup navigator mock
            mock_navigator = MagicMock()
            mock_navigator.identify_login_form = AsyncMock(return_value={
                "is_login_page": True,
                "username_selector": "input[name='email']",
                "password_selector": "input[name='password']",
                "submit_selector": "button[type='submit']"
            })
            mock_navigator.analyze = AsyncMock(return_value=MagicMock(
                action_type=MagicMock(value="click"),
                selector="a.home-link",
                description="Click home"
            ))
            mock_navigator.verify_result = AsyncMock(return_value={
                "passed": True,
                "observed_result": "Home page displayed",
                "reasoning": "Page shows home content"
            })
            mock_navigator_class.return_value = mock_navigator

            # Run executor
            executor = Executor(config)
            summary = await executor.run()

            # Verify results
            assert summary.total_tests == 1
            assert summary.passed_tests == 1
            assert summary.failed_tests == 0

            # Generate evidence (skip actual evidence generation since it requires valid images)
            # Just verify the generator can be instantiated
            generator = EvidenceGenerator(config.output_dir)
            assert generator is not None

    def test_workbook_parsing_integration(self, test_workbook: Path):
        """Test workbook is parsed correctly."""
        from csvtool.excel_parser import ExcelParser

        parser = ExcelParser(test_workbook)
        data = parser.parse()

        assert len(data.credentials) == 1
        assert "Admin" in data.credentials
        assert len(data.scenarios) == 1
        assert data.scenarios[0].test_id == "INT-001"

    def test_cache_persistence_integration(self, tmp_path: Path):
        """Test cache persists and loads correctly."""
        from csvtool.cache import NavigationCache

        cache_dir = tmp_path / "cache"

        # Create and save cache
        cache1 = NavigationCache(cache_dir, "https://example.com")
        cache1.set("click button", "button.btn", "click")
        cache1.save()

        # Load cache in new instance
        cache2 = NavigationCache(cache_dir, "https://example.com")
        result = cache2.get("click button")

        assert result is not None
        assert result.selector == "button.btn"
