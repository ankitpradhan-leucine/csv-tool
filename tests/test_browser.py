"""Tests for browser automation wrapper."""

import pytest
from pathlib import Path
from unittest.mock import Mock, AsyncMock, patch
from csvtool.browser import Browser, BrowserConfig


class TestBrowserConfig:
    """Tests for BrowserConfig."""

    def test_default_config(self):
        """Test default browser configuration."""
        config = BrowserConfig()
        assert config.headless is True
        assert config.timeout == 30000
        assert config.screenshot_dir == Path("output/screenshots")

    def test_custom_config(self):
        """Test custom browser configuration."""
        config = BrowserConfig(
            headless=False,
            timeout=60000,
            screenshot_dir=Path("/custom/path")
        )
        assert config.headless is False
        assert config.timeout == 60000


class TestBrowser:
    """Tests for Browser class."""

    @pytest.fixture
    def browser_config(self, tmp_path: Path) -> BrowserConfig:
        """Create browser config with temp screenshot dir."""
        return BrowserConfig(
            headless=True,
            screenshot_dir=tmp_path / "screenshots"
        )

    @pytest.mark.asyncio
    async def test_browser_context_manager(self, browser_config: BrowserConfig):
        """Test browser can be used as async context manager."""
        async with Browser(browser_config) as browser:
            assert browser.page is not None

    @pytest.mark.asyncio
    async def test_navigate_to_url(self, browser_config: BrowserConfig):
        """Test navigating to a URL."""
        async with Browser(browser_config) as browser:
            await browser.navigate("https://example.com")
            assert "example.com" in browser.page.url

    @pytest.mark.asyncio
    async def test_take_screenshot(self, browser_config: BrowserConfig):
        """Test taking a screenshot."""
        async with Browser(browser_config) as browser:
            await browser.navigate("https://example.com")
            screenshot_path = await browser.screenshot("test_screenshot")

            assert screenshot_path.exists()
            assert screenshot_path.suffix == ".png"

    @pytest.mark.asyncio
    async def test_screenshot_includes_timestamp(self, browser_config: BrowserConfig):
        """Test screenshot filename includes timestamp."""
        async with Browser(browser_config) as browser:
            await browser.navigate("https://example.com")
            screenshot_path = await browser.screenshot("step_1")

            # Filename format: step_1_20260313_143000.png
            assert "step_1_" in screenshot_path.name

    @pytest.mark.asyncio
    async def test_get_page_content(self, browser_config: BrowserConfig):
        """Test getting page HTML content."""
        async with Browser(browser_config) as browser:
            await browser.navigate("https://example.com")
            content = await browser.get_page_content()

            assert "<html" in content.lower()

    @pytest.mark.asyncio
    async def test_click_element(self, browser_config: BrowserConfig):
        """Test clicking an element by selector."""
        async with Browser(browser_config) as browser:
            await browser.navigate("https://example.com")
            # Example.com has a link, try to click it
            await browser.click("a")

    @pytest.mark.asyncio
    async def test_type_text(self, browser_config: BrowserConfig):
        """Test typing text into an input field."""
        async with Browser(browser_config) as browser:
            await browser.navigate("https://example.com")
            # This will fail on example.com but tests the interface
            with pytest.raises(Exception):
                await browser.type_text("input[name='email']", "test@example.com")

    @pytest.mark.asyncio
    async def test_get_current_url(self, browser_config: BrowserConfig):
        """Test getting current URL."""
        async with Browser(browser_config) as browser:
            await browser.navigate("https://example.com")
            url = browser.get_current_url()
            assert "example.com" in url

    @pytest.mark.asyncio
    async def test_wait_for_selector(self, browser_config: BrowserConfig):
        """Test waiting for an element to appear."""
        async with Browser(browser_config) as browser:
            await browser.navigate("https://example.com")
            # Example.com has an h1 element
            result = await browser.wait_for_selector("h1")
            assert result is True

            # Non-existent selector should return False
            result = await browser.wait_for_selector("nonexistent", timeout=1000)
            assert result is False
