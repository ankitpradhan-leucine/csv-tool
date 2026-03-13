"""Playwright browser automation wrapper."""

import base64
from datetime import datetime
from pathlib import Path
from typing import Optional

from playwright.async_api import (
    TimeoutError as PlaywrightTimeoutError,
    async_playwright,
    Browser as PWBrowser,
    Page,
)
from pydantic import BaseModel


class BrowserConfig(BaseModel):
    """Configuration for browser automation."""

    headless: bool = True
    timeout: int = 30000  # milliseconds
    screenshot_dir: Path = Path("output/screenshots")
    viewport_width: int = 1920
    viewport_height: int = 1080


class Browser:
    """Async browser automation wrapper using Playwright."""

    def __init__(self, config: Optional[BrowserConfig] = None):
        """Initialize browser with configuration."""
        self.config = config or BrowserConfig()
        self._playwright = None
        self._browser: Optional[PWBrowser] = None
        self._page: Optional[Page] = None

        # Ensure screenshot directory exists
        self.config.screenshot_dir.mkdir(parents=True, exist_ok=True)

    async def __aenter__(self) -> "Browser":
        """Async context manager entry - launch browser."""
        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(
            headless=self.config.headless
        )
        self._page = await self._browser.new_page(
            viewport={
                "width": self.config.viewport_width,
                "height": self.config.viewport_height
            }
        )
        self._page.set_default_timeout(self.config.timeout)
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """Async context manager exit - close browser."""
        if self._page:
            await self._page.close()
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()

    @property
    def page(self) -> Page:
        """Get current page, raising if not initialized."""
        if self._page is None:
            raise RuntimeError("Browser not initialized. Use 'async with Browser()' context.")
        return self._page

    async def navigate(self, url: str) -> None:
        """Navigate to a URL and wait for load."""
        await self.page.goto(url, wait_until="networkidle")

    async def screenshot(self, name: str) -> Path:
        """Take a full-page screenshot and return path."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{name}_{timestamp}.png"
        path = self.config.screenshot_dir / filename

        await self.page.screenshot(path=str(path), full_page=True)
        return path

    async def get_page_content(self) -> str:
        """Get current page HTML content."""
        return await self.page.content()

    async def get_screenshot_base64(self) -> str:
        """Get screenshot as base64 string for AI analysis."""
        screenshot_bytes = await self.page.screenshot(full_page=True)
        return base64.b64encode(screenshot_bytes).decode("utf-8")

    async def click(self, selector: str) -> None:
        """Click an element by selector."""
        await self.page.click(selector)
        await self.page.wait_for_load_state("networkidle")

    async def type_text(self, selector: str, text: str) -> None:
        """Type text into an input field."""
        await self.page.fill(selector, text)

    async def wait_for_selector(self, selector: str, timeout: Optional[int] = None) -> bool:
        """Wait for an element to appear."""
        try:
            await self.page.wait_for_selector(
                selector,
                timeout=timeout or self.config.timeout
            )
            return True
        except PlaywrightTimeoutError:
            return False

    def get_current_url(self) -> str:
        """Get current page URL."""
        return self.page.url

    async def clear_session(self) -> None:
        """Clear cookies and storage for new session."""
        context = self.page.context
        await context.clear_cookies()
        await self.page.evaluate("window.localStorage.clear()")
        await self.page.evaluate("window.sessionStorage.clear()")
