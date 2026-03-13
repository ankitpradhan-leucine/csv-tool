"""Tests for test executor."""

import pytest
from pathlib import Path
from datetime import datetime
from unittest.mock import Mock, AsyncMock, MagicMock, patch

from csvtool.executor import Executor, ExecutorConfig
from csvtool.models import (
    Credential, TestScenario, TestStep, StepResult,
    TestResult, ExecutionSummary
)
from csvtool.excel_parser import WorkbookData


class TestExecutorConfig:
    """Tests for ExecutorConfig."""

    def test_default_config(self):
        """Test default executor configuration."""
        config = ExecutorConfig(
            url="https://app.example.com",
            workbook_path=Path("test.xlsx")
        )
        assert config.url == "https://app.example.com"
        assert config.stop_on_failure is True

    def test_config_with_options(self):
        """Test executor config with options."""
        config = ExecutorConfig(
            url="https://app.example.com",
            workbook_path=Path("test.xlsx"),
            stop_on_failure=False,
            headless=False
        )
        assert config.stop_on_failure is False
        assert config.headless is False


class TestExecutor:
    """Tests for Executor class."""

    @pytest.fixture
    def mock_workbook_data(self) -> WorkbookData:
        """Create mock workbook data."""
        return WorkbookData(
            credentials={
                "Admin": Credential(
                    role="Admin",
                    username="admin@test.com",
                    password="admin123"
                )
            },
            scenarios=[
                TestScenario(
                    test_id="TEST-001",
                    role="Admin",
                    title="Test scenario",
                    instructions="1. Navigate to settings\n2. Verify page loads",
                    expected_result="Settings page displayed"
                )
            ],
            workbook_name="test.xlsx"
        )

    @pytest.fixture
    def executor_config(self, tmp_path: Path) -> ExecutorConfig:
        """Create executor config."""
        return ExecutorConfig(
            url="https://app.example.com",
            workbook_path=tmp_path / "test.xlsx",
            output_dir=tmp_path / "output",
            cache_dir=tmp_path / "cache"
        )

    @pytest.fixture
    def mock_browser(self):
        """Create mock browser."""
        browser = AsyncMock()
        browser.navigate = AsyncMock()
        browser.screenshot = AsyncMock(return_value=Path("/tmp/screenshot.png"))
        browser.get_screenshot_base64 = AsyncMock(return_value="base64data")
        browser.click = AsyncMock()
        browser.type_text = AsyncMock()
        browser.get_current_url = Mock(return_value="https://app.example.com")
        browser.clear_session = AsyncMock()
        return browser

    @pytest.fixture
    def mock_navigator(self):
        """Create mock AI navigator."""
        navigator = AsyncMock()
        navigator.identify_login_form = AsyncMock(return_value={
            "username_selector": "input[name='email']",
            "password_selector": "input[name='password']",
            "submit_selector": "button[type='submit']",
            "is_login_page": True
        })
        navigator.analyze = AsyncMock(return_value=MagicMock(
            action_type="click",
            selector="button.settings",
            description="Click settings"
        ))
        navigator.verify_result = AsyncMock(return_value={
            "passed": True,
            "observed_result": "Settings page displayed",
            "reasoning": "Page shows settings"
        })
        return navigator

    @pytest.mark.asyncio
    async def test_executor_initialization(
        self,
        executor_config: ExecutorConfig,
        mock_workbook_data: WorkbookData
    ):
        """Test executor initializes correctly."""
        with patch("csvtool.executor.ExcelParser") as mock_parser:
            mock_parser.return_value.parse.return_value = mock_workbook_data

            executor = Executor(executor_config)
            assert executor is not None

    @pytest.mark.asyncio
    async def test_execute_step_uses_cache(
        self,
        executor_config: ExecutorConfig,
        mock_workbook_data: WorkbookData,
        mock_browser,
        mock_navigator
    ):
        """Test that executor uses cached selectors when available."""
        with patch("csvtool.executor.ExcelParser") as mock_parser:
            mock_parser.return_value.parse.return_value = mock_workbook_data

            executor = Executor(executor_config)
            executor._browser = mock_browser
            executor._navigator = mock_navigator

            # Pre-populate cache
            executor._cache.set(
                "navigate to settings",
                "a.settings-link",
                "click"
            )

            step = TestStep(number=1, instruction="Navigate to settings")
            await executor._execute_step(step)

            # Should use cached selector, not call navigator
            mock_navigator.analyze.assert_not_called()
            mock_browser.click.assert_called_with("a.settings-link")

    @pytest.mark.asyncio
    async def test_execute_step_falls_back_to_ai(
        self,
        executor_config: ExecutorConfig,
        mock_workbook_data: WorkbookData,
        mock_browser,
        mock_navigator
    ):
        """Test that executor falls back to AI when cache misses."""
        with patch("csvtool.executor.ExcelParser") as mock_parser:
            mock_parser.return_value.parse.return_value = mock_workbook_data

            executor = Executor(executor_config)
            executor._browser = mock_browser
            executor._navigator = mock_navigator

            step = TestStep(number=1, instruction="Click the submit button")
            await executor._execute_step(step)

            # Should call navigator since not in cache
            mock_navigator.analyze.assert_called_once()
