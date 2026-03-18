"""Test execution orchestrator."""

import asyncio
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)

from csvtool.ai_navigator import AINavigator, ActionType
from csvtool.browser import Browser, BrowserConfig
from csvtool.cache import NavigationCache
from csvtool.excel_parser import ExcelParser, WorkbookData
from csvtool.models import (
    Credential, ExecutionSummary, StepResult,
    TestResult, TestScenario, TestStep
)


class ExecutorConfig(BaseModel):
    """Configuration for test executor."""

    url: str
    workbook_path: Path
    output_dir: Path = Path("output")
    cache_dir: Path = Path("cache")
    headless: bool = True
    stop_on_failure: bool = True
    api_key: Optional[str] = None


class ExecutionError(Exception):
    """Error during test execution."""
    pass


class Executor:
    """Orchestrates test scenario execution."""

    def __init__(self, config: ExecutorConfig):
        """Initialize executor with configuration."""
        self.config = config

        # Parse workbook
        parser = ExcelParser(config.workbook_path)
        self._workbook_data = parser.parse()

        # Initialize components (lazy)
        self._browser: Optional[Browser] = None
        self._navigator: Optional[AINavigator] = None
        self._cache = NavigationCache(
            cache_dir=config.cache_dir,
            product_url=config.url
        )

        # Execution state
        self._current_screenshots: list[Path] = []
        self._step_results: list[StepResult] = []

    async def run(self) -> ExecutionSummary:
        """Run all test scenarios and return summary."""
        browser_config = BrowserConfig(
            headless=self.config.headless,
            screenshot_dir=self.config.output_dir / "screenshots"
        )

        self._navigator = AINavigator(api_key=self.config.api_key)

        test_results: list[TestResult] = []

        async with Browser(browser_config) as browser:
            self._browser = browser

            for scenario in self._workbook_data.scenarios:
                # Reset token counters for per-scenario cost tracking
                tokens_before_input = self._navigator.total_input_tokens
                tokens_before_output = self._navigator.total_output_tokens

                try:
                    result = await self._execute_scenario(scenario)
                    test_results.append(result)

                    # Calculate cost for this scenario
                    self._log_scenario_cost(
                        scenario.test_id,
                        self._navigator.total_input_tokens - tokens_before_input,
                        self._navigator.total_output_tokens - tokens_before_output
                    )

                    if not result.passed and self.config.stop_on_failure:
                        break

                except Exception as e:
                    # Create failed result
                    result = TestResult(
                        test_id=scenario.test_id,
                        title=scenario.title,
                        role=scenario.role,
                        passed=False,
                        expected_result=scenario.expected_result,
                        observed_result="Execution error",
                        step_results=self._step_results.copy(),
                        error=str(e)
                    )
                    test_results.append(result)

                    # Log cost even for failed scenarios
                    self._log_scenario_cost(
                        scenario.test_id,
                        self._navigator.total_input_tokens - tokens_before_input,
                        self._navigator.total_output_tokens - tokens_before_output
                    )

                    if self.config.stop_on_failure:
                        break

        # Save cache after run
        self._cache.save()

        # Log total cost for all scenarios
        total_input = self._navigator.total_input_tokens
        total_output = self._navigator.total_output_tokens
        total_input_cost = (total_input / 1_000_000) * 3.0
        total_output_cost = (total_output / 1_000_000) * 15.0
        total_cost = total_input_cost + total_output_cost

        logger.info("")
        logger.info("=" * 80)
        logger.info(
            f"💰 TOTAL COST ESTIMATION | "
            f"Input: {total_input} tokens (${total_input_cost:.4f}) | "
            f"Output: {total_output} tokens (${total_output_cost:.4f}) | "
            f"Total: {total_input + total_output} tokens (${total_cost:.4f})"
        )
        logger.info("=" * 80)

        # Build summary
        passed = sum(1 for r in test_results if r.passed)
        return ExecutionSummary(
            url=self.config.url,
            workbook_name=self._workbook_data.workbook_name,
            execution_date=datetime.now(),
            total_tests=len(test_results),
            passed_tests=passed,
            failed_tests=len(test_results) - passed,
            test_results=test_results
        )

    def _log_scenario_cost(self, test_id: str, input_tokens: int, output_tokens: int) -> None:
        """Log cost estimation for a test scenario."""
        # Claude Sonnet 4 pricing (as of 2025)
        # Input: $3 per million tokens
        # Output: $15 per million tokens
        input_cost = (input_tokens / 1_000_000) * 3.0
        output_cost = (output_tokens / 1_000_000) * 15.0
        total_cost = input_cost + output_cost

        logger.info(
            f"💰 Cost Estimation - {test_id} | "
            f"Input: {input_tokens} tokens (${input_cost:.4f}) | "
            f"Output: {output_tokens} tokens (${output_cost:.4f}) | "
            f"Total: {input_tokens + output_tokens} tokens (${total_cost:.4f})"
        )

    async def _execute_scenario(self, scenario: TestScenario) -> TestResult:
        """Execute a single test scenario."""
        self._step_results = []

        # Get credentials for role
        credential = self._workbook_data.get_credential(scenario.role)

        # Navigate to URL and capture login page (before entering credentials)
        await self._browser.navigate(self.config.url)

        # Capture login page - this is what the user requested to see
        initial_screenshot = await self._take_screenshot(f"{scenario.test_id}_login_page")
        self._step_results.append(StepResult(
            step_number=0,
            instruction=f"Login Page: {self.config.url} (before entering credentials as {credential.role})",
            screenshot_path=initial_screenshot,
            timestamp=datetime.now(),
            success=True
        ))

        # Perform login
        await self._login(credential)

        # No need for separate post-login screenshot - the first test step will show that

        # Execute each step
        for step in scenario.steps:
            step_result = await self._execute_step(step, scenario.test_id)
            self._step_results.append(step_result)

            if not step_result.success:
                return TestResult(
                    test_id=scenario.test_id,
                    title=scenario.title,
                    role=scenario.role,
                    passed=False,
                    expected_result=scenario.expected_result,
                    observed_result=step_result.error or "Step failed",
                    step_results=self._step_results,
                    error=step_result.error
                )

        # Verify expected result
        screenshot_b64 = await self._browser.get_screenshot_base64()
        verification = await self._navigator.verify_result(
            screenshot_base64=screenshot_b64,
            expected_result=scenario.expected_result
        )

        await self._take_screenshot(f"{scenario.test_id}_final")

        # Clear session for next test
        await self._browser.clear_session()

        return TestResult(
            test_id=scenario.test_id,
            title=scenario.title,
            role=scenario.role,
            passed=verification.get("passed", False),
            expected_result=scenario.expected_result,
            observed_result=verification.get("observed_result", ""),
            step_results=self._step_results,
            error=None if verification.get("passed") else verification.get("reasoning")
        )

    async def _login(self, credential: Credential) -> None:
        """Perform login with given credentials."""
        screenshot_b64 = await self._browser.get_screenshot_base64()

        login_info = await self._navigator.identify_login_form(screenshot_b64)

        if not login_info.get("is_login_page"):
            # Not a login page, might already be logged in
            return

        # Enter credentials
        await self._browser.type_text(
            login_info["username_selector"],
            credential.username
        )
        await self._browser.type_text(
            login_info["password_selector"],
            credential.password
        )

        # Submit
        await self._browser.click(login_info["submit_selector"])

        # Wait for navigation
        await asyncio.sleep(2)

    async def _execute_step(
        self,
        step: TestStep,
        test_id: str = ""
    ) -> StepResult:
        """Execute a single test step."""
        screenshot_name = f"{test_id}_step_{step.number}"

        try:
            # Check cache first
            cached = self._cache.get(step.instruction)

            if cached:
                # Use cached selector
                await self._perform_action(cached.action, cached.selector)
                self._cache.set(step.instruction, cached.selector, cached.action)
            else:
                # Fall back to AI
                screenshot_b64 = await self._browser.get_screenshot_base64()
                action = await self._navigator.analyze(
                    screenshot_base64=screenshot_b64,
                    instruction=step.instruction
                )

                await self._perform_action(
                    action.action_type.value,
                    action.selector,
                    action.text
                )

                # Cache the learned selector
                if action.selector:
                    self._cache.set(
                        step.instruction,
                        action.selector,
                        action.action_type.value
                    )

            # Take screenshot after step
            screenshot_path = await self._take_screenshot(screenshot_name)

            return StepResult(
                step_number=step.number,
                instruction=step.instruction,
                screenshot_path=screenshot_path,
                timestamp=datetime.now(),
                success=True
            )

        except Exception as e:
            # Take error screenshot
            screenshot_path = await self._take_screenshot(f"{screenshot_name}_error")

            return StepResult(
                step_number=step.number,
                instruction=step.instruction,
                screenshot_path=screenshot_path,
                timestamp=datetime.now(),
                success=False,
                error=str(e)
            )

    async def _perform_action(
        self,
        action_type: str,
        selector: Optional[str] = None,
        text: Optional[str] = None
    ) -> None:
        """Perform a browser action."""
        if action_type == "click" and selector:
            await self._browser.click(selector)
        elif action_type == "type" and selector and text:
            await self._browser.type_text(selector, text)
        elif action_type == "navigate" and text:
            await self._browser.navigate(text)
        elif action_type == "wait":
            await asyncio.sleep(2)
        # verify actions are handled separately

    async def _take_screenshot(self, name: str) -> Path:
        """Take a screenshot and track it."""
        path = await self._browser.screenshot(name)
        self._current_screenshots.append(path)
        return path
