# CSV Automation Tool Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a tool that automates CSV test execution across web applications using Playwright and Claude AI, generating Word document evidence.

**Architecture:** CLI tool reads test scenarios from Excel workbooks, uses Playwright for browser automation with Claude AI for intelligent navigation, captures screenshots at every step, and generates Word documents with full traceability. Navigation cache reduces API costs on subsequent runs.

**Tech Stack:** Python 3.11+, Playwright, Anthropic Claude API, openpyxl, python-docx, Click, Pydantic

**Spec Reference:** `docs/superpowers/specs/2026-03-13-csv-automation-tool-design.md`

---

## File Structure

```
csv-automation/
├── src/
│   └── csvtool/
│       ├── __init__.py           # Package init, version
│       ├── main.py               # CLI entry point (Click)
│       ├── models.py             # Pydantic data models
│       ├── excel_parser.py       # Workbook loading and parsing
│       ├── browser.py            # Playwright wrapper
│       ├── ai_navigator.py       # Claude API integration
│       ├── cache.py              # Navigation cache management
│       ├── executor.py           # Test execution orchestration
│       └── evidence.py           # Word document generation
├── tests/
│   ├── __init__.py
│   ├── conftest.py               # Shared fixtures
│   ├── test_models.py
│   ├── test_excel_parser.py
│   ├── test_browser.py
│   ├── test_ai_navigator.py
│   ├── test_cache.py
│   ├── test_executor.py
│   ├── test_evidence.py
│   └── test_integration.py
├── scenarios/                    # Test workbooks (user-provided)
├── cache/                        # Navigation caches (generated)
├── output/                       # Evidence documents (generated)
├── .github/
│   └── workflows/
│       └── run-tests.yml         # GitHub Actions workflow
├── pyproject.toml                # Project config, dependencies
├── requirements.txt              # Pinned dependencies
└── README.md
```

---

## Chunk 1: Project Setup and Core Models

### Task 1.1: Initialize Project Structure

**Files:**
- Create: `pyproject.toml`
- Create: `requirements.txt`
- Create: `src/csvtool/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`

- [ ] **Step 1: Create pyproject.toml**

```toml
[build-system]
requires = ["setuptools>=61.0"]
build-backend = "setuptools.build_meta"

[project]
name = "csvtool"
version = "0.1.0"
description = "CSV test automation tool with AI-powered navigation"
readme = "README.md"
requires-python = ">=3.11"
dependencies = [
    "playwright>=1.40.0",
    "anthropic>=0.18.0",
    "openpyxl>=3.1.0",
    "python-docx>=1.1.0",
    "click>=8.1.0",
    "pydantic>=2.0.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=7.4.0",
    "pytest-asyncio>=0.21.0",
    "pytest-mock>=3.12.0",
]

[project.scripts]
csvtool = "csvtool.main:cli"

[tool.setuptools.packages.find]
where = ["src"]
```

- [ ] **Step 2: Create requirements.txt**

```
playwright>=1.40.0
anthropic>=0.18.0
openpyxl>=3.1.0
python-docx>=1.1.0
click>=8.1.0
pydantic>=2.0.0
pytest>=7.4.0
pytest-asyncio>=0.21.0
pytest-mock>=3.12.0
```

- [ ] **Step 3: Create src/csvtool/__init__.py**

```python
"""CSV Automation Tool - Automated test execution with AI-powered navigation."""

__version__ = "0.1.0"
```

- [ ] **Step 4: Create tests/__init__.py**

```python
"""Test suite for CSV Automation Tool."""
```

- [ ] **Step 5: Create tests/conftest.py**

```python
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
```

- [ ] **Step 6: Create directories**

```bash
mkdir -p src/csvtool tests/fixtures scenarios cache output .github/workflows
```

- [ ] **Step 7: Install dependencies and verify**

```bash
pip install -e ".[dev]"
playwright install chromium
python -c "import csvtool; print(csvtool.__version__)"
```

Expected: `0.1.0`

- [ ] **Step 8: Commit**

```bash
git add pyproject.toml requirements.txt src/ tests/
git commit -m "chore: initialize project structure with dependencies"
```

---

### Task 1.2: Define Core Data Models

**Files:**
- Create: `src/csvtool/models.py`
- Create: `tests/test_models.py`

- [ ] **Step 1: Write failing tests for Credential model**

```python
# tests/test_models.py
"""Tests for data models."""

import pytest
from csvtool.models import Credential, TestScenario, TestStep, StepResult, TestResult


class TestCredential:
    """Tests for Credential model."""

    def test_credential_creation(self):
        """Test creating a credential with all fields."""
        cred = Credential(
            role="Admin",
            username="admin@example.com",
            password="secret123"
        )
        assert cred.role == "Admin"
        assert cred.username == "admin@example.com"
        assert cred.password == "secret123"

    def test_credential_role_required(self):
        """Test that role is required."""
        with pytest.raises(ValueError):
            Credential(role="", username="user", password="pass")
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_models.py::TestCredential -v
```

Expected: FAIL with "cannot import name 'Credential'"

- [ ] **Step 3: Write failing tests for TestStep model**

```python
# Append to tests/test_models.py

class TestTestStep:
    """Tests for TestStep model."""

    def test_step_creation(self):
        """Test creating a test step."""
        step = TestStep(
            number=1,
            instruction="Open the settings menu"
        )
        assert step.number == 1
        assert step.instruction == "Open the settings menu"

    def test_step_from_instruction_text(self):
        """Test parsing step from numbered instruction."""
        step = TestStep.from_text("1. Open the settings menu")
        assert step.number == 1
        assert step.instruction == "Open the settings menu"

    def test_step_from_instruction_strips_whitespace(self):
        """Test that instruction whitespace is stripped."""
        step = TestStep.from_text("  2.  Click on Audit Log  ")
        assert step.number == 2
        assert step.instruction == "Click on Audit Log"
```

- [ ] **Step 4: Write failing tests for TestScenario model**

```python
# Append to tests/test_models.py

class TestTestScenario:
    """Tests for TestScenario model."""

    def test_scenario_creation(self):
        """Test creating a test scenario."""
        scenario = TestScenario(
            test_id="URAP-01",
            role="Quality Specialist",
            title="Creates protocol with empty field",
            instructions="1. Open form\n2. Leave field blank\n3. Submit",
            expected_result="Error message shown"
        )
        assert scenario.test_id == "URAP-01"
        assert scenario.role == "Quality Specialist"
        assert len(scenario.steps) == 3

    def test_scenario_parses_steps(self):
        """Test that instructions are parsed into steps."""
        scenario = TestScenario(
            test_id="TEST-01",
            role="Admin",
            title="Test",
            instructions="1. Step one\n2. Step two",
            expected_result="Success"
        )
        assert scenario.steps[0].instruction == "Step one"
        assert scenario.steps[1].instruction == "Step two"

    def test_scenario_handles_multiline_steps(self):
        """Test parsing instructions with various formats."""
        scenario = TestScenario(
            test_id="TEST-01",
            role="Admin",
            title="Test",
            instructions="1. First step\n\n2. Second step\n3. Third step",
            expected_result="Success"
        )
        assert len(scenario.steps) == 3
```

- [ ] **Step 5: Write failing tests for result models**

```python
# Append to tests/test_models.py

from pathlib import Path
from datetime import datetime


class TestStepResult:
    """Tests for StepResult model."""

    def test_step_result_creation(self):
        """Test creating a step result."""
        result = StepResult(
            step_number=1,
            instruction="Open form",
            screenshot_path=Path("/tmp/screenshot.png"),
            timestamp=datetime(2026, 3, 13, 14, 30, 0),
            success=True
        )
        assert result.step_number == 1
        assert result.success is True

    def test_step_result_with_error(self):
        """Test step result with error."""
        result = StepResult(
            step_number=2,
            instruction="Click button",
            screenshot_path=Path("/tmp/error.png"),
            timestamp=datetime(2026, 3, 13, 14, 30, 5),
            success=False,
            error="Element not found"
        )
        assert result.success is False
        assert result.error == "Element not found"


class TestTestResult:
    """Tests for TestResult model."""

    def test_result_passed(self):
        """Test passed result."""
        result = TestResult(
            test_id="TEST-01",
            title="Test scenario",
            role="Admin",
            passed=True,
            expected_result="Success",
            observed_result="Success",
            step_results=[]
        )
        assert result.passed is True

    def test_result_failed(self):
        """Test failed result."""
        result = TestResult(
            test_id="TEST-01",
            title="Test scenario",
            role="Admin",
            passed=False,
            expected_result="Success",
            observed_result="Error occurred",
            step_results=[],
            error="Verification failed"
        )
        assert result.passed is False
        assert result.error == "Verification failed"
```

- [ ] **Step 6: Run all model tests to verify they fail**

```bash
pytest tests/test_models.py -v
```

Expected: Multiple FAILures with import errors

- [ ] **Step 7: Implement models**

```python
# src/csvtool/models.py
"""Data models for CSV Automation Tool."""

import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from pydantic import BaseModel, Field, field_validator


class Credential(BaseModel):
    """User credential for a specific role."""

    role: str = Field(..., min_length=1)
    username: str = Field(..., min_length=1)
    password: str = Field(..., min_length=1)

    @field_validator("role")
    @classmethod
    def role_not_empty(cls, v: str) -> str:
        if not v.strip():
            raise ValueError("Role cannot be empty")
        return v.strip()


class TestStep(BaseModel):
    """A single step in a test scenario."""

    number: int = Field(..., ge=1)
    instruction: str = Field(..., min_length=1)

    @classmethod
    def from_text(cls, text: str) -> "TestStep":
        """Parse a step from numbered instruction text like '1. Do something'."""
        text = text.strip()
        match = re.match(r"(\d+)\.\s*(.+)", text)
        if not match:
            raise ValueError(f"Invalid step format: {text}")
        return cls(number=int(match.group(1)), instruction=match.group(2).strip())


class TestScenario(BaseModel):
    """A test scenario with multiple steps."""

    test_id: str = Field(..., min_length=1)
    role: str = Field(..., min_length=1)
    title: str = Field(..., min_length=1)
    instructions: str = Field(..., min_length=1)
    expected_result: str = Field(..., min_length=1)
    steps: list[TestStep] = Field(default_factory=list)

    def model_post_init(self, __context) -> None:
        """Parse instructions into steps after initialization."""
        if not self.steps:
            self.steps = self._parse_steps(self.instructions)

    @staticmethod
    def _parse_steps(instructions: str) -> list[TestStep]:
        """Parse numbered instructions into TestStep objects."""
        steps = []
        lines = instructions.strip().split("\n")
        for line in lines:
            line = line.strip()
            if line and re.match(r"\d+\.", line):
                steps.append(TestStep.from_text(line))
        return steps


class StepResult(BaseModel):
    """Result of executing a single test step."""

    step_number: int
    instruction: str
    screenshot_path: Path
    timestamp: datetime
    success: bool
    error: Optional[str] = None


class TestResult(BaseModel):
    """Result of executing a complete test scenario."""

    test_id: str
    title: str
    role: str
    passed: bool
    expected_result: str
    observed_result: str
    step_results: list[StepResult]
    error: Optional[str] = None


class ExecutionSummary(BaseModel):
    """Summary of a complete test run."""

    url: str
    workbook_name: str
    execution_date: datetime
    total_tests: int
    passed_tests: int
    failed_tests: int
    test_results: list[TestResult]

    @property
    def success_rate(self) -> float:
        """Calculate success rate as percentage."""
        if self.total_tests == 0:
            return 0.0
        return (self.passed_tests / self.total_tests) * 100
```

- [ ] **Step 8: Run tests to verify they pass**

```bash
pytest tests/test_models.py -v
```

Expected: All PASS

- [ ] **Step 9: Commit**

```bash
git add src/csvtool/models.py tests/test_models.py
git commit -m "feat: add core data models with Pydantic validation"
```

---

## Chunk 2: Excel Parser

### Task 2.1: Implement Excel Parser

**Files:**
- Create: `src/csvtool/excel_parser.py`
- Create: `tests/test_excel_parser.py`
- Create: `tests/fixtures/sample_workbook.xlsx`

- [ ] **Step 1: Create sample test workbook fixture**

```python
# Run this script to create the fixture
# tests/fixtures/create_sample_workbook.py

from openpyxl import Workbook

wb = Workbook()

# Credentials sheet
creds_sheet = wb.active
creds_sheet.title = "Credentials"
creds_sheet.append(["Role", "Username", "Password"])
creds_sheet.append(["Admin", "admin@example.com", "admin123"])
creds_sheet.append(["Quality Specialist", "qs@example.com", "qs123"])
creds_sheet.append(["Operator", "operator@example.com", "op123"])

# Test Scenarios sheet
scenarios_sheet = wb.create_sheet("Test Scenarios")
scenarios_sheet.append(["Test ID", "Role", "Title", "Test Instructions", "Expected Result"])
scenarios_sheet.append([
    "TEST-001",
    "Admin",
    "Verify admin can view settings",
    "1. Navigate to Settings\n2. Verify page loads",
    "Settings page displayed"
])
scenarios_sheet.append([
    "TEST-002",
    "Quality Specialist",
    "Create protocol with empty field",
    "1. Open protocol form\n2. Leave mandatory field blank\n3. Click submit",
    "Error message shown"
])

wb.save("tests/fixtures/sample_workbook.xlsx")
print("Created sample_workbook.xlsx")
```

```bash
mkdir -p tests/fixtures
python tests/fixtures/create_sample_workbook.py
```

- [ ] **Step 2: Write failing tests for excel parser**

```python
# tests/test_excel_parser.py
"""Tests for Excel parser."""

import pytest
from pathlib import Path
from csvtool.excel_parser import ExcelParser, WorkbookData
from csvtool.models import Credential, TestScenario


class TestExcelParser:
    """Tests for ExcelParser."""

    def test_load_workbook(self, sample_workbook_path: Path):
        """Test loading a workbook."""
        parser = ExcelParser(sample_workbook_path)
        data = parser.parse()
        assert isinstance(data, WorkbookData)

    def test_parse_credentials(self, sample_workbook_path: Path):
        """Test parsing credentials sheet."""
        parser = ExcelParser(sample_workbook_path)
        data = parser.parse()

        assert len(data.credentials) == 3
        assert data.credentials["Admin"].username == "admin@example.com"
        assert data.credentials["Quality Specialist"].username == "qs@example.com"

    def test_get_credential_by_role(self, sample_workbook_path: Path):
        """Test getting credential by role name."""
        parser = ExcelParser(sample_workbook_path)
        data = parser.parse()

        cred = data.get_credential("Admin")
        assert cred.username == "admin@example.com"
        assert cred.password == "admin123"

    def test_get_credential_invalid_role(self, sample_workbook_path: Path):
        """Test getting credential for non-existent role."""
        parser = ExcelParser(sample_workbook_path)
        data = parser.parse()

        with pytest.raises(KeyError):
            data.get_credential("NonExistentRole")

    def test_parse_scenarios(self, sample_workbook_path: Path):
        """Test parsing test scenarios."""
        parser = ExcelParser(sample_workbook_path)
        data = parser.parse()

        assert len(data.scenarios) == 2
        assert data.scenarios[0].test_id == "TEST-001"
        assert data.scenarios[0].role == "Admin"

    def test_scenario_steps_parsed(self, sample_workbook_path: Path):
        """Test that scenario instructions are parsed into steps."""
        parser = ExcelParser(sample_workbook_path)
        data = parser.parse()

        scenario = data.scenarios[1]
        assert len(scenario.steps) == 3
        assert scenario.steps[0].instruction == "Open protocol form"

    def test_workbook_not_found(self):
        """Test error when workbook doesn't exist."""
        with pytest.raises(FileNotFoundError):
            parser = ExcelParser(Path("/nonexistent/workbook.xlsx"))
            parser.parse()

    def test_missing_credentials_sheet(self, tmp_path: Path):
        """Test error when Credentials sheet is missing."""
        from openpyxl import Workbook

        wb = Workbook()
        wb.active.title = "Other"
        path = tmp_path / "bad_workbook.xlsx"
        wb.save(path)

        parser = ExcelParser(path)
        with pytest.raises(ValueError, match="Credentials"):
            parser.parse()
```

- [ ] **Step 3: Run tests to verify they fail**

```bash
pytest tests/test_excel_parser.py -v
```

Expected: FAIL with import error

- [ ] **Step 4: Implement ExcelParser**

```python
# src/csvtool/excel_parser.py
"""Excel workbook parser for test scenarios and credentials."""

from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from csvtool.models import Credential, TestScenario


class WorkbookData(BaseModel):
    """Parsed workbook data containing credentials and scenarios."""

    credentials: dict[str, Credential]
    scenarios: list[TestScenario]
    workbook_name: str

    def get_credential(self, role: str) -> Credential:
        """Get credential by role name."""
        if role not in self.credentials:
            raise KeyError(f"No credential found for role: {role}")
        return self.credentials[role]


class ExcelParser:
    """Parser for Excel workbooks containing test scenarios."""

    CREDENTIALS_SHEET = "Credentials"
    CREDENTIALS_HEADERS = ["Role", "Username", "Password"]
    SCENARIO_HEADERS = ["Test ID", "Role", "Title", "Test Instructions", "Expected Result"]

    def __init__(self, workbook_path: Path):
        """Initialize parser with workbook path."""
        self.workbook_path = Path(workbook_path)
        if not self.workbook_path.exists():
            raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    def parse(self) -> WorkbookData:
        """Parse the workbook and return structured data."""
        wb = load_workbook(self.workbook_path, read_only=True, data_only=True)

        try:
            credentials = self._parse_credentials(wb)
            scenarios = self._parse_scenarios(wb)

            return WorkbookData(
                credentials=credentials,
                scenarios=scenarios,
                workbook_name=self.workbook_path.name
            )
        finally:
            wb.close()

    def _parse_credentials(self, wb) -> dict[str, Credential]:
        """Parse the Credentials sheet."""
        if self.CREDENTIALS_SHEET not in wb.sheetnames:
            raise ValueError(
                f"Workbook must contain a '{self.CREDENTIALS_SHEET}' sheet"
            )

        sheet = wb[self.CREDENTIALS_SHEET]
        credentials = {}

        rows = list(sheet.iter_rows(min_row=2, values_only=True))
        for row in rows:
            if row[0] is None:
                continue
            role, username, password = row[0], row[1], row[2]
            credentials[role] = Credential(
                role=str(role),
                username=str(username),
                password=str(password)
            )

        return credentials

    def _parse_scenarios(self, wb) -> list[TestScenario]:
        """Parse all scenario sheets (any sheet that's not Credentials)."""
        scenarios = []

        for sheet_name in wb.sheetnames:
            if sheet_name == self.CREDENTIALS_SHEET:
                continue

            sheet = wb[sheet_name]
            scenarios.extend(self._parse_scenario_sheet(sheet))

        return scenarios

    def _parse_scenario_sheet(self, sheet: Worksheet) -> list[TestScenario]:
        """Parse a single scenario sheet."""
        scenarios = []

        rows = list(sheet.iter_rows(min_row=2, values_only=True))
        for row in rows:
            if row[0] is None:
                continue

            test_id = str(row[0]) if row[0] else ""
            role = str(row[1]) if row[1] else ""
            title = str(row[2]) if row[2] else ""
            instructions = str(row[3]) if row[3] else ""
            expected_result = str(row[4]) if row[4] else ""

            if test_id and instructions:
                scenarios.append(TestScenario(
                    test_id=test_id,
                    role=role,
                    title=title,
                    instructions=instructions,
                    expected_result=expected_result
                ))

        return scenarios
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
pytest tests/test_excel_parser.py -v
```

Expected: All PASS

- [ ] **Step 6: Commit**

```bash
git add src/csvtool/excel_parser.py tests/test_excel_parser.py tests/fixtures/
git commit -m "feat: add Excel parser for workbooks with credentials and scenarios"
```

---

## Chunk 3: Browser Automation

### Task 3.1: Implement Playwright Browser Wrapper

**Files:**
- Create: `src/csvtool/browser.py`
- Create: `tests/test_browser.py`

- [ ] **Step 1: Write failing tests for Browser class**

```python
# tests/test_browser.py
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
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_browser.py -v
```

Expected: FAIL with import error

- [ ] **Step 3: Implement Browser class**

```python
# src/csvtool/browser.py
"""Playwright browser automation wrapper."""

import asyncio
from datetime import datetime
from pathlib import Path
from typing import Optional

from playwright.async_api import async_playwright, Browser as PWBrowser, Page
from pydantic import BaseModel, Field


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
        import base64
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
        except Exception:
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
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_browser.py -v
```

Expected: All PASS (some may be skipped if no network)

- [ ] **Step 5: Commit**

```bash
git add src/csvtool/browser.py tests/test_browser.py
git commit -m "feat: add Playwright browser wrapper with screenshot support"
```

---

## Chunk 4: AI Navigator

### Task 4.1: Implement Claude AI Navigator

**Files:**
- Create: `src/csvtool/ai_navigator.py`
- Create: `tests/test_ai_navigator.py`

- [ ] **Step 1: Write failing tests for AINavigator**

```python
# tests/test_ai_navigator.py
"""Tests for AI navigator using Claude."""

import pytest
from unittest.mock import Mock, AsyncMock, patch, MagicMock
from csvtool.ai_navigator import AINavigator, NavigationAction, ActionType


class TestNavigationAction:
    """Tests for NavigationAction model."""

    def test_click_action(self):
        """Test creating a click action."""
        action = NavigationAction(
            action_type=ActionType.CLICK,
            selector="button.submit",
            description="Click submit button"
        )
        assert action.action_type == ActionType.CLICK
        assert action.selector == "button.submit"

    def test_type_action(self):
        """Test creating a type action."""
        action = NavigationAction(
            action_type=ActionType.TYPE,
            selector="input[name='email']",
            text="test@example.com",
            description="Enter email"
        )
        assert action.action_type == ActionType.TYPE
        assert action.text == "test@example.com"

    def test_verify_action(self):
        """Test creating a verify action."""
        action = NavigationAction(
            action_type=ActionType.VERIFY,
            expected_text="Success",
            description="Verify success message"
        )
        assert action.action_type == ActionType.VERIFY


class TestAINavigator:
    """Tests for AINavigator class."""

    @pytest.fixture
    def mock_anthropic_client(self):
        """Create mock Anthropic client."""
        with patch("csvtool.ai_navigator.anthropic.Anthropic") as mock:
            client = MagicMock()
            mock.return_value = client
            yield client

    @pytest.fixture
    def navigator(self, mock_anthropic_client):
        """Create AINavigator with mocked client."""
        return AINavigator(api_key="test-key")

    def test_navigator_initialization(self, mock_anthropic_client):
        """Test navigator initializes with API key."""
        navigator = AINavigator(api_key="test-key")
        assert navigator is not None

    @pytest.mark.asyncio
    async def test_analyze_returns_action(self, navigator, mock_anthropic_client):
        """Test analyze returns a NavigationAction."""
        # Mock Claude response
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text='''
        {
            "action_type": "click",
            "selector": "button.login",
            "description": "Click login button"
        }
        ''')]
        mock_anthropic_client.messages.create.return_value = mock_response

        action = await navigator.analyze(
            screenshot_base64="base64data",
            instruction="Click the login button"
        )

        assert isinstance(action, NavigationAction)
        assert action.action_type == ActionType.CLICK

    @pytest.mark.asyncio
    async def test_verify_result(self, navigator, mock_anthropic_client):
        """Test verifying expected result against screenshot."""
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text='''
        {
            "passed": true,
            "observed_result": "Success message displayed",
            "reasoning": "The page shows 'Operation successful'"
        }
        ''')]
        mock_anthropic_client.messages.create.return_value = mock_response

        result = await navigator.verify_result(
            screenshot_base64="base64data",
            expected_result="Success message displayed"
        )

        assert result["passed"] is True
        assert "observed_result" in result

    @pytest.mark.asyncio
    async def test_generate_test_data(self, navigator, mock_anthropic_client):
        """Test generating test data for form fields."""
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text='''
        {
            "fields": [
                {"selector": "input[name='name']", "value": "John Doe"},
                {"selector": "input[name='email']", "value": "john@example.com"}
            ]
        }
        ''')]
        mock_anthropic_client.messages.create.return_value = mock_response

        data = await navigator.generate_test_data(
            screenshot_base64="base64data",
            instruction="Enter details",
            skip_mandatory=False
        )

        assert len(data["fields"]) == 2
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_ai_navigator.py -v
```

Expected: FAIL with import error

- [ ] **Step 3: Implement AINavigator**

```python
# src/csvtool/ai_navigator.py
"""AI-powered navigation using Claude API."""

import json
import os
from enum import Enum
from typing import Optional

import anthropic
from pydantic import BaseModel, Field


class ActionType(str, Enum):
    """Types of browser actions."""
    CLICK = "click"
    TYPE = "type"
    VERIFY = "verify"
    WAIT = "wait"
    NAVIGATE = "navigate"


class NavigationAction(BaseModel):
    """An action to perform in the browser."""

    action_type: ActionType
    selector: Optional[str] = None
    text: Optional[str] = None
    expected_text: Optional[str] = None
    url: Optional[str] = None
    description: str = ""


class AINavigator:
    """AI-powered navigation using Claude Vision API."""

    MODEL = "claude-sonnet-4-20250514"
    MAX_TOKENS = 1024

    def __init__(self, api_key: Optional[str] = None):
        """Initialize with Anthropic API key."""
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not self.api_key:
            raise ValueError("ANTHROPIC_API_KEY must be provided or set in environment")
        self.client = anthropic.Anthropic(api_key=self.api_key)

    async def analyze(
        self,
        screenshot_base64: str,
        instruction: str,
        context: Optional[str] = None
    ) -> NavigationAction:
        """Analyze screenshot and determine next action for instruction."""
        system_prompt = """You are a browser automation assistant. Analyze the screenshot and determine the exact action needed to fulfill the instruction.

Return a JSON object with:
- action_type: "click", "type", "verify", "wait", or "navigate"
- selector: CSS selector for the element (for click/type actions)
- text: text to type (for type actions)
- expected_text: text to verify (for verify actions)
- description: brief description of the action

Return ONLY valid JSON, no markdown or explanation."""

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": screenshot_base64
                }
            },
            {
                "type": "text",
                "text": f"Instruction: {instruction}\n\nContext: {context or 'None'}\n\nWhat action should be taken?"
            }
        ]

        response = self.client.messages.create(
            model=self.MODEL,
            max_tokens=self.MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )

        response_text = response.content[0].text.strip()

        # Parse JSON response
        try:
            # Handle potential markdown code blocks
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]

            data = json.loads(response_text)
            return NavigationAction(
                action_type=ActionType(data.get("action_type", "click")),
                selector=data.get("selector"),
                text=data.get("text"),
                expected_text=data.get("expected_text"),
                description=data.get("description", "")
            )
        except (json.JSONDecodeError, KeyError) as e:
            raise ValueError(f"Failed to parse AI response: {response_text}") from e

    async def verify_result(
        self,
        screenshot_base64: str,
        expected_result: str
    ) -> dict:
        """Verify if the expected result matches the current screen state."""
        system_prompt = """You are a test verification assistant. Compare the screenshot against the expected result.

Return a JSON object with:
- passed: true/false
- observed_result: what you actually see on the screen
- reasoning: brief explanation of your verification

Return ONLY valid JSON, no markdown or explanation."""

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": screenshot_base64
                }
            },
            {
                "type": "text",
                "text": f"Expected Result: {expected_result}\n\nDoes the screenshot match the expected result?"
            }
        ]

        response = self.client.messages.create(
            model=self.MODEL,
            max_tokens=self.MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )

        response_text = response.content[0].text.strip()

        try:
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]

            return json.loads(response_text)
        except json.JSONDecodeError as e:
            raise ValueError(f"Failed to parse verification response: {response_text}") from e

    async def generate_test_data(
        self,
        screenshot_base64: str,
        instruction: str,
        skip_mandatory: bool = False
    ) -> dict:
        """Generate appropriate test data for form fields visible on screen."""
        skip_instruction = ""
        if skip_mandatory:
            skip_instruction = "IMPORTANT: Identify one mandatory/required field and leave it empty (return empty string for its value)."

        system_prompt = f"""You are a test data generator. Analyze the form fields visible in the screenshot and generate appropriate test data.

{skip_instruction}

Return a JSON object with:
- fields: array of objects with "selector" (CSS selector) and "value" (data to enter)

Generate realistic dummy data appropriate for each field type (names, emails, dates, numbers, etc.).
Return ONLY valid JSON, no markdown or explanation."""

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": screenshot_base64
                }
            },
            {
                "type": "text",
                "text": f"Instruction: {instruction}\n\nGenerate test data for the visible form fields."
            }
        ]

        response = self.client.messages.create(
            model=self.MODEL,
            max_tokens=self.MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )

        response_text = response.content[0].text.strip()

        try:
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]

            return json.loads(response_text)
        except json.JSONDecodeError as e:
            raise ValueError(f"Failed to parse test data response: {response_text}") from e

    async def identify_login_form(self, screenshot_base64: str) -> dict:
        """Identify login form fields on the page."""
        system_prompt = """You are a login form analyzer. Identify the username/email and password fields on the login page.

Return a JSON object with:
- username_selector: CSS selector for username/email field
- password_selector: CSS selector for password field
- submit_selector: CSS selector for login/submit button
- is_login_page: true/false

Return ONLY valid JSON, no markdown or explanation."""

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": screenshot_base64
                }
            },
            {
                "type": "text",
                "text": "Identify the login form elements on this page."
            }
        ]

        response = self.client.messages.create(
            model=self.MODEL,
            max_tokens=self.MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )

        response_text = response.content[0].text.strip()

        try:
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]

            return json.loads(response_text)
        except json.JSONDecodeError as e:
            raise ValueError(f"Failed to parse login form response: {response_text}") from e
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_ai_navigator.py -v
```

Expected: All PASS

- [ ] **Step 5: Commit**

```bash
git add src/csvtool/ai_navigator.py tests/test_ai_navigator.py
git commit -m "feat: add Claude AI navigator for intelligent browser navigation"
```

---

## Chunk 5: Navigation Cache

### Task 5.1: Implement Navigation Cache

**Files:**
- Create: `src/csvtool/cache.py`
- Create: `tests/test_cache.py`

- [ ] **Step 1: Write failing tests for NavigationCache**

```python
# tests/test_cache.py
"""Tests for navigation cache."""

import pytest
import json
from pathlib import Path
from datetime import datetime
from csvtool.cache import NavigationCache, CachedNavigation


class TestCachedNavigation:
    """Tests for CachedNavigation model."""

    def test_cached_navigation_creation(self):
        """Test creating a cached navigation entry."""
        cached = CachedNavigation(
            selector="button.submit",
            action="click",
            success_count=5,
            last_success=datetime.now()
        )
        assert cached.selector == "button.submit"
        assert cached.success_count == 5

    def test_increment_success(self):
        """Test incrementing success count."""
        cached = CachedNavigation(
            selector="button.submit",
            action="click",
            success_count=5,
            last_success=datetime.now()
        )
        cached.record_success()
        assert cached.success_count == 6


class TestNavigationCache:
    """Tests for NavigationCache."""

    @pytest.fixture
    def cache_dir(self, tmp_path: Path) -> Path:
        """Create temporary cache directory."""
        cache_path = tmp_path / "cache"
        cache_path.mkdir()
        return cache_path

    @pytest.fixture
    def cache(self, cache_dir: Path) -> NavigationCache:
        """Create cache instance."""
        return NavigationCache(
            cache_dir=cache_dir,
            product_url="https://app.example.com"
        )

    def test_cache_initialization(self, cache: NavigationCache):
        """Test cache initializes empty."""
        assert len(cache.navigations) == 0

    def test_cache_get_miss(self, cache: NavigationCache):
        """Test cache miss returns None."""
        result = cache.get("nonexistent instruction")
        assert result is None

    def test_cache_set_and_get(self, cache: NavigationCache):
        """Test setting and getting cache entry."""
        cache.set(
            instruction="click login button",
            selector="button.login",
            action="click"
        )

        result = cache.get("click login button")
        assert result is not None
        assert result.selector == "button.login"

    def test_cache_case_insensitive(self, cache: NavigationCache):
        """Test cache lookup is case insensitive."""
        cache.set(
            instruction="Click Login Button",
            selector="button.login",
            action="click"
        )

        result = cache.get("click login button")
        assert result is not None

    def test_cache_normalizes_whitespace(self, cache: NavigationCache):
        """Test cache normalizes whitespace in instructions."""
        cache.set(
            instruction="  click   login   button  ",
            selector="button.login",
            action="click"
        )

        result = cache.get("click login button")
        assert result is not None

    def test_cache_persists_to_file(self, cache: NavigationCache, cache_dir: Path):
        """Test cache saves to JSON file."""
        cache.set(
            instruction="click login",
            selector="button.login",
            action="click"
        )
        cache.save()

        cache_file = cache_dir / "app.example.com.json"
        assert cache_file.exists()

        data = json.loads(cache_file.read_text())
        assert "click login" in data["navigations"]

    def test_cache_loads_from_file(self, cache_dir: Path):
        """Test cache loads existing file."""
        # Create cache file
        cache_file = cache_dir / "app.example.com.json"
        cache_data = {
            "cache_version": "1.0",
            "product_url": "https://app.example.com",
            "last_updated": datetime.now().isoformat(),
            "navigations": {
                "click submit": {
                    "selector": "button[type='submit']",
                    "action": "click",
                    "success_count": 10,
                    "last_success": datetime.now().isoformat()
                }
            }
        }
        cache_file.write_text(json.dumps(cache_data))

        # Load cache
        cache = NavigationCache(
            cache_dir=cache_dir,
            product_url="https://app.example.com"
        )

        result = cache.get("click submit")
        assert result is not None
        assert result.success_count == 10

    def test_cache_update_increments_count(self, cache: NavigationCache):
        """Test updating existing cache entry."""
        cache.set("click button", "button.btn", "click")
        cache.set("click button", "button.btn", "click")

        result = cache.get("click button")
        assert result.success_count == 2

    def test_cache_update_with_new_selector(self, cache: NavigationCache):
        """Test updating cache with different selector."""
        cache.set("click button", "button.old", "click")
        cache.set("click button", "button.new", "click")

        result = cache.get("click button")
        assert result.selector == "button.new"
        assert result.success_count == 1  # Reset on selector change
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_cache.py -v
```

Expected: FAIL with import error

- [ ] **Step 3: Implement NavigationCache**

```python
# src/csvtool/cache.py
"""Navigation cache for storing learned selectors."""

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

from pydantic import BaseModel, Field


class CachedNavigation(BaseModel):
    """A cached navigation entry."""

    selector: str
    action: str
    success_count: int = 1
    last_success: datetime = Field(default_factory=datetime.now)

    def record_success(self) -> None:
        """Record a successful use of this cached navigation."""
        self.success_count += 1
        self.last_success = datetime.now()


class NavigationCache:
    """Cache for storing learned navigation selectors."""

    CACHE_VERSION = "1.0"

    def __init__(self, cache_dir: Path, product_url: str):
        """Initialize cache for a specific product URL."""
        self.cache_dir = Path(cache_dir)
        self.product_url = product_url
        self.navigations: dict[str, CachedNavigation] = {}

        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self._load()

    @property
    def cache_file(self) -> Path:
        """Get cache file path based on product URL."""
        # Extract domain from URL for filename
        parsed = urlparse(self.product_url)
        domain = parsed.netloc or parsed.path
        # Sanitize for filename
        safe_name = re.sub(r"[^\w\-.]", "_", domain)
        return self.cache_dir / f"{safe_name}.json"

    @staticmethod
    def _normalize_instruction(instruction: str) -> str:
        """Normalize instruction for cache key."""
        # Lowercase, collapse whitespace, strip
        normalized = instruction.lower().strip()
        normalized = re.sub(r"\s+", " ", normalized)
        return normalized

    def get(self, instruction: str) -> Optional[CachedNavigation]:
        """Get cached navigation for instruction."""
        key = self._normalize_instruction(instruction)
        return self.navigations.get(key)

    def set(
        self,
        instruction: str,
        selector: str,
        action: str
    ) -> None:
        """Set or update cached navigation."""
        key = self._normalize_instruction(instruction)

        existing = self.navigations.get(key)
        if existing and existing.selector == selector:
            # Same selector - increment success count
            existing.record_success()
        else:
            # New entry or different selector
            self.navigations[key] = CachedNavigation(
                selector=selector,
                action=action,
                success_count=1,
                last_success=datetime.now()
            )

    def save(self) -> None:
        """Save cache to file."""
        data = {
            "cache_version": self.CACHE_VERSION,
            "product_url": self.product_url,
            "last_updated": datetime.now().isoformat(),
            "navigations": {
                key: {
                    "selector": nav.selector,
                    "action": nav.action,
                    "success_count": nav.success_count,
                    "last_success": nav.last_success.isoformat()
                }
                for key, nav in self.navigations.items()
            }
        }
        self.cache_file.write_text(json.dumps(data, indent=2))

    def _load(self) -> None:
        """Load cache from file if exists."""
        if not self.cache_file.exists():
            return

        try:
            data = json.loads(self.cache_file.read_text())

            for key, nav_data in data.get("navigations", {}).items():
                self.navigations[key] = CachedNavigation(
                    selector=nav_data["selector"],
                    action=nav_data["action"],
                    success_count=nav_data.get("success_count", 1),
                    last_success=datetime.fromisoformat(
                        nav_data.get("last_success", datetime.now().isoformat())
                    )
                )
        except (json.JSONDecodeError, KeyError) as e:
            # Invalid cache file - start fresh
            self.navigations = {}
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_cache.py -v
```

Expected: All PASS

- [ ] **Step 5: Commit**

```bash
git add src/csvtool/cache.py tests/test_cache.py
git commit -m "feat: add navigation cache with persistence"
```

---

## Chunk 6: Test Executor

### Task 6.1: Implement Test Executor

**Files:**
- Create: `src/csvtool/executor.py`
- Create: `tests/test_executor.py`

- [ ] **Step 1: Write failing tests for Executor**

```python
# tests/test_executor.py
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
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_executor.py -v
```

Expected: FAIL with import error

- [ ] **Step 3: Implement Executor**

```python
# src/csvtool/executor.py
"""Test execution orchestrator."""

import asyncio
from datetime import datetime
from pathlib import Path
from typing import Optional

from pydantic import BaseModel, Field

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
                try:
                    result = await self._execute_scenario(scenario)
                    test_results.append(result)

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

                    if self.config.stop_on_failure:
                        break

        # Save cache after run
        self._cache.save()

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

    async def _execute_scenario(self, scenario: TestScenario) -> TestResult:
        """Execute a single test scenario."""
        self._step_results = []

        # Get credentials for role
        credential = self._workbook_data.get_credential(scenario.role)

        # Navigate to URL and login
        await self._browser.navigate(self.config.url)
        await self._take_screenshot(f"{scenario.test_id}_initial")

        await self._login(credential)
        await self._take_screenshot(f"{scenario.test_id}_logged_in")

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
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_executor.py -v
```

Expected: All PASS

- [ ] **Step 5: Commit**

```bash
git add src/csvtool/executor.py tests/test_executor.py
git commit -m "feat: add test executor with cache integration"
```

---

## Chunk 7: Evidence Generator

### Task 7.1: Implement Word Document Evidence Generator

**Files:**
- Create: `src/csvtool/evidence.py`
- Create: `tests/test_evidence.py`

- [ ] **Step 1: Write failing tests for EvidenceGenerator**

```python
# tests/test_evidence.py
"""Tests for evidence document generator."""

import pytest
from pathlib import Path
from datetime import datetime
from docx import Document

from csvtool.evidence import EvidenceGenerator
from csvtool.models import (
    ExecutionSummary, TestResult, StepResult
)


class TestEvidenceGenerator:
    """Tests for EvidenceGenerator."""

    @pytest.fixture
    def output_dir(self, tmp_path: Path) -> Path:
        """Create temporary output directory."""
        output = tmp_path / "output"
        output.mkdir()
        return output

    @pytest.fixture
    def sample_screenshot(self, tmp_path: Path) -> Path:
        """Create a sample screenshot file."""
        # Create a simple PNG file
        from PIL import Image
        img = Image.new("RGB", (100, 100), color="red")
        path = tmp_path / "screenshot.png"
        img.save(path)
        return path

    @pytest.fixture
    def sample_summary(self, sample_screenshot: Path) -> ExecutionSummary:
        """Create sample execution summary."""
        return ExecutionSummary(
            url="https://app.example.com",
            workbook_name="test_workbook.xlsx",
            execution_date=datetime(2026, 3, 13, 14, 30, 0),
            total_tests=2,
            passed_tests=1,
            failed_tests=1,
            test_results=[
                TestResult(
                    test_id="TEST-001",
                    title="Verify admin access",
                    role="Admin",
                    passed=True,
                    expected_result="Settings page displayed",
                    observed_result="Settings page displayed",
                    step_results=[
                        StepResult(
                            step_number=1,
                            instruction="Navigate to settings",
                            screenshot_path=sample_screenshot,
                            timestamp=datetime(2026, 3, 13, 14, 30, 5),
                            success=True
                        ),
                        StepResult(
                            step_number=2,
                            instruction="Verify page loads",
                            screenshot_path=sample_screenshot,
                            timestamp=datetime(2026, 3, 13, 14, 30, 10),
                            success=True
                        )
                    ]
                ),
                TestResult(
                    test_id="TEST-002",
                    title="Create protocol",
                    role="Quality Specialist",
                    passed=False,
                    expected_result="Protocol created",
                    observed_result="Error: Field required",
                    step_results=[
                        StepResult(
                            step_number=1,
                            instruction="Open form",
                            screenshot_path=sample_screenshot,
                            timestamp=datetime(2026, 3, 13, 14, 31, 0),
                            success=True
                        )
                    ],
                    error="Validation failed"
                )
            ]
        )

    def test_generator_creates_document(
        self,
        output_dir: Path,
        sample_summary: ExecutionSummary
    ):
        """Test generator creates a Word document."""
        generator = EvidenceGenerator(output_dir)
        doc_path = generator.generate(sample_summary)

        assert doc_path.exists()
        assert doc_path.suffix == ".docx"

    def test_document_has_header(
        self,
        output_dir: Path,
        sample_summary: ExecutionSummary
    ):
        """Test document contains header with run info."""
        generator = EvidenceGenerator(output_dir)
        doc_path = generator.generate(sample_summary)

        doc = Document(doc_path)
        text = "\n".join([p.text for p in doc.paragraphs])

        assert "CSV TEST EVIDENCE REPORT" in text
        assert "https://app.example.com" in text
        assert "test_workbook.xlsx" in text

    def test_document_has_summary(
        self,
        output_dir: Path,
        sample_summary: ExecutionSummary
    ):
        """Test document contains test summary."""
        generator = EvidenceGenerator(output_dir)
        doc_path = generator.generate(sample_summary)

        doc = Document(doc_path)
        text = "\n".join([p.text for p in doc.paragraphs])

        assert "Total Tests: 2" in text or "Total: 2" in text
        assert "Passed: 1" in text
        assert "Failed: 1" in text

    def test_document_has_test_results(
        self,
        output_dir: Path,
        sample_summary: ExecutionSummary
    ):
        """Test document contains individual test results."""
        generator = EvidenceGenerator(output_dir)
        doc_path = generator.generate(sample_summary)

        doc = Document(doc_path)
        text = "\n".join([p.text for p in doc.paragraphs])

        assert "TEST-001" in text
        assert "TEST-002" in text
        assert "Verify admin access" in text
        assert "PASSED" in text
        assert "FAILED" in text

    def test_document_has_screenshots(
        self,
        output_dir: Path,
        sample_summary: ExecutionSummary
    ):
        """Test document contains embedded screenshots."""
        generator = EvidenceGenerator(output_dir)
        doc_path = generator.generate(sample_summary)

        doc = Document(doc_path)

        # Check for inline shapes (images)
        image_count = 0
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run._element.xpath('.//a:blip'):
                    image_count += 1

        # Also check in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if run._element.xpath('.//a:blip'):
                                image_count += 1

        # Should have at least some images
        assert image_count > 0 or len(doc.inline_shapes) > 0

    def test_document_filename_format(
        self,
        output_dir: Path,
        sample_summary: ExecutionSummary
    ):
        """Test document filename follows expected format."""
        generator = EvidenceGenerator(output_dir)
        doc_path = generator.generate(sample_summary)

        # Format: CSV-Evidence-{workbook}-{date}-{time}.docx
        assert "CSV-Evidence" in doc_path.name
        assert "test_workbook" in doc_path.name
```

- [ ] **Step 2: Add Pillow to dependencies for test fixture**

```bash
pip install Pillow
```

- [ ] **Step 3: Run tests to verify they fail**

```bash
pytest tests/test_evidence.py -v
```

Expected: FAIL with import error

- [ ] **Step 4: Implement EvidenceGenerator**

```python
# src/csvtool/evidence.py
"""Word document evidence generator."""

from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

from csvtool.models import ExecutionSummary, TestResult, StepResult


class EvidenceGenerator:
    """Generates Word document evidence reports."""

    def __init__(self, output_dir: Path):
        """Initialize generator with output directory."""
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def generate(self, summary: ExecutionSummary) -> Path:
        """Generate evidence document from execution summary."""
        doc = Document()

        # Set up styles
        self._setup_styles(doc)

        # Add header
        self._add_header(doc, summary)

        # Add summary table
        self._add_summary_table(doc, summary)

        # Add each test result
        for result in summary.test_results:
            self._add_test_result(doc, result)

        # Save document
        filename = self._generate_filename(summary)
        doc_path = self.output_dir / filename
        doc.save(str(doc_path))

        return doc_path

    def _setup_styles(self, doc: Document) -> None:
        """Set up document styles."""
        # Modify Normal style
        style = doc.styles["Normal"]
        style.font.name = "Arial"
        style.font.size = Pt(11)

    def _add_header(self, doc: Document, summary: ExecutionSummary) -> None:
        """Add document header."""
        # Title
        title = doc.add_heading("CSV TEST EVIDENCE REPORT", 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Metadata
        doc.add_paragraph()
        meta = doc.add_paragraph()
        meta.add_run(f"Application URL: ").bold = True
        meta.add_run(summary.url)

        meta = doc.add_paragraph()
        meta.add_run(f"Workbook: ").bold = True
        meta.add_run(summary.workbook_name)

        meta = doc.add_paragraph()
        meta.add_run(f"Execution Date: ").bold = True
        meta.add_run(summary.execution_date.strftime("%Y-%m-%d %H:%M:%S"))

        doc.add_paragraph()

    def _add_summary_table(self, doc: Document, summary: ExecutionSummary) -> None:
        """Add summary statistics table."""
        doc.add_heading("Execution Summary", level=1)

        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"

        # Header row
        header_cells = table.rows[0].cells
        header_cells[0].text = "Total Tests"
        header_cells[1].text = "Passed"
        header_cells[2].text = "Failed"
        header_cells[3].text = "Success Rate"

        # Make headers bold
        for cell in header_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True

        # Data row
        row = table.add_row().cells
        row[0].text = str(summary.total_tests)
        row[1].text = str(summary.passed_tests)
        row[2].text = str(summary.failed_tests)
        row[3].text = f"{summary.success_rate:.1f}%"

        doc.add_paragraph()

    def _add_test_result(self, doc: Document, result: TestResult) -> None:
        """Add a single test result section."""
        # Test header
        status = "PASSED" if result.passed else "FAILED"
        status_color = "green" if result.passed else "red"

        doc.add_heading(f"TEST: {result.test_id}", level=2)

        # Test metadata table
        table = doc.add_table(rows=4, cols=2)
        table.style = "Table Grid"

        rows = table.rows
        rows[0].cells[0].text = "Title"
        rows[0].cells[1].text = result.title
        rows[1].cells[0].text = "Role"
        rows[1].cells[1].text = result.role
        rows[2].cells[0].text = "Status"
        rows[2].cells[1].text = status
        rows[3].cells[0].text = "Expected Result"
        rows[3].cells[1].text = result.expected_result

        # Make first column bold
        for row in rows:
            for paragraph in row.cells[0].paragraphs:
                for run in paragraph.runs:
                    run.bold = True

        doc.add_paragraph()

        # Steps
        doc.add_heading("Steps", level=3)

        for step_result in result.step_results:
            self._add_step_result(doc, step_result)

        # Verification
        doc.add_heading("Verification", level=3)

        verify_table = doc.add_table(rows=2, cols=2)
        verify_table.style = "Table Grid"

        verify_table.rows[0].cells[0].text = "Expected Result"
        verify_table.rows[0].cells[1].text = result.expected_result
        verify_table.rows[1].cells[0].text = "Observed Result"
        verify_table.rows[1].cells[1].text = result.observed_result

        if result.error:
            error_row = verify_table.add_row().cells
            error_row[0].text = "Error"
            error_row[1].text = result.error

        doc.add_paragraph()
        doc.add_paragraph("─" * 50)  # Separator
        doc.add_paragraph()

    def _add_step_result(self, doc: Document, step: StepResult) -> None:
        """Add a single step result with screenshot."""
        # Step header
        status_icon = "✓" if step.success else "✗"
        para = doc.add_paragraph()
        para.add_run(f"Step {step.number}: ").bold = True
        para.add_run(step.instruction)
        para.add_run(f" {status_icon}")

        # Timestamp
        time_para = doc.add_paragraph()
        time_para.add_run("Timestamp: ").italic = True
        time_para.add_run(step.timestamp.strftime("%H:%M:%S"))

        # Screenshot
        if step.screenshot_path and step.screenshot_path.exists():
            try:
                doc.add_picture(
                    str(step.screenshot_path),
                    width=Inches(6)
                )
                # Center the image
                last_paragraph = doc.paragraphs[-1]
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception as e:
                doc.add_paragraph(f"[Screenshot unavailable: {e}]")

        # Error message if failed
        if not step.success and step.error:
            error_para = doc.add_paragraph()
            error_para.add_run("Error: ").bold = True
            error_para.add_run(step.error)

        doc.add_paragraph()

    def _generate_filename(self, summary: ExecutionSummary) -> str:
        """Generate filename for evidence document."""
        # Remove extension from workbook name
        workbook_base = Path(summary.workbook_name).stem
        date_str = summary.execution_date.strftime("%Y%m%d_%H%M%S")
        return f"CSV-Evidence-{workbook_base}-{date_str}.docx"
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
pytest tests/test_evidence.py -v
```

Expected: All PASS

- [ ] **Step 6: Commit**

```bash
git add src/csvtool/evidence.py tests/test_evidence.py
git commit -m "feat: add Word document evidence generator"
```

---

## Chunk 8: CLI Interface

### Task 8.1: Implement CLI with Click

**Files:**
- Create: `src/csvtool/main.py`
- Create: `tests/test_main.py`

- [ ] **Step 1: Write failing tests for CLI**

```python
# tests/test_main.py
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
                total_tests=1, passed_tests=1, failed_tests=0
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
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_main.py -v
```

Expected: FAIL with import error

- [ ] **Step 3: Implement CLI**

```python
# src/csvtool/main.py
"""CLI entry point for CSV Automation Tool."""

import asyncio
import sys
from pathlib import Path
from typing import Optional

import click

from csvtool import __version__
from csvtool.evidence import EvidenceGenerator
from csvtool.executor import Executor, ExecutorConfig


@click.group()
@click.version_option(version=__version__, prog_name="CSV Automation Tool")
def cli():
    """CSV Automation Tool - Automated test execution with AI-powered navigation."""
    pass


@cli.command()
@click.option(
    "--url",
    required=True,
    help="Application URL to test"
)
@click.option(
    "--workbook",
    required=True,
    type=click.Path(exists=True, path_type=Path),
    help="Path to test scenarios Excel workbook"
)
@click.option(
    "--output",
    type=click.Path(path_type=Path),
    default=Path("output"),
    help="Output directory for evidence documents"
)
@click.option(
    "--cache-dir",
    type=click.Path(path_type=Path),
    default=Path("cache"),
    help="Directory for navigation cache"
)
@click.option(
    "--headless/--no-headless",
    default=True,
    help="Run browser in headless mode"
)
@click.option(
    "--stop-on-failure/--continue-on-failure",
    default=True,
    help="Stop execution on first failure"
)
@click.option(
    "--api-key",
    envvar="ANTHROPIC_API_KEY",
    help="Anthropic API key (or set ANTHROPIC_API_KEY env var)"
)
def run(
    url: str,
    workbook: Path,
    output: Path,
    cache_dir: Path,
    headless: bool,
    stop_on_failure: bool,
    api_key: Optional[str]
):
    """Run CSV test scenarios against an application."""
    click.echo(f"CSV Automation Tool v{__version__}")
    click.echo(f"URL: {url}")
    click.echo(f"Workbook: {workbook}")
    click.echo()

    # Create config
    config = ExecutorConfig(
        url=url,
        workbook_path=workbook,
        output_dir=output,
        cache_dir=cache_dir,
        headless=headless,
        stop_on_failure=stop_on_failure,
        api_key=api_key
    )

    # Run executor
    click.echo("Starting test execution...")
    executor = Executor(config)

    try:
        summary = asyncio.run(executor.run())
    except Exception as e:
        click.echo(f"Execution error: {e}", err=True)
        sys.exit(1)

    # Generate evidence
    click.echo("Generating evidence document...")
    generator = EvidenceGenerator(output)
    doc_path = generator.generate(summary)

    # Print summary
    click.echo()
    click.echo("=" * 50)
    click.echo("EXECUTION SUMMARY")
    click.echo("=" * 50)
    click.echo(f"Total Tests: {summary.total_tests}")
    click.echo(f"Passed: {summary.passed_tests}")
    click.echo(f"Failed: {summary.failed_tests}")
    click.echo(f"Success Rate: {summary.success_rate:.1f}%")
    click.echo()
    click.echo(f"Evidence Document: {doc_path}")

    # Exit with appropriate code
    if summary.failed_tests > 0:
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    cli()
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_main.py -v
```

Expected: All PASS

- [ ] **Step 5: Test CLI manually**

```bash
csvtool --help
csvtool --version
csvtool run --help
```

- [ ] **Step 6: Commit**

```bash
git add src/csvtool/main.py tests/test_main.py
git commit -m "feat: add CLI interface with Click"
```

---

## Chunk 9: GitHub Actions Workflow

### Task 9.1: Create GitHub Actions Workflow

**Files:**
- Create: `.github/workflows/run-tests.yml`
- Create: `README.md`

- [ ] **Step 1: Create GitHub Actions workflow**

```yaml
# .github/workflows/run-tests.yml
name: Run CSV Tests

on:
  workflow_dispatch:
    inputs:
      application_url:
        description: 'Application URL to test'
        required: true
        type: string

      workbook:
        description: 'Test workbook file'
        required: true
        type: choice
        options:
          - scenarios/workbook-1.xlsx
          - scenarios/workbook-2.xlsx
          - scenarios/workbook-3.xlsx

env:
  PYTHON_VERSION: '3.11'

jobs:
  run-csv-tests:
    runs-on: self-hosted
    timeout-minutes: 120

    steps:
      - name: Checkout repository
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: ${{ env.PYTHON_VERSION }}
          cache: 'pip'

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -e .
          playwright install chromium
          playwright install-deps chromium

      - name: Run CSV Tests
        env:
          ANTHROPIC_API_KEY: ${{ secrets.ANTHROPIC_API_KEY }}
        run: |
          csvtool run \
            --url "${{ inputs.application_url }}" \
            --workbook "${{ inputs.workbook }}" \
            --output output

      - name: Upload Evidence
        uses: actions/upload-artifact@v4
        if: always()
        with:
          name: csv-evidence-${{ github.run_number }}
          path: output/*.docx
          retention-days: 90

      - name: Upload Screenshots
        uses: actions/upload-artifact@v4
        if: always()
        with:
          name: csv-screenshots-${{ github.run_number }}
          path: output/screenshots/
          retention-days: 30
```

- [ ] **Step 2: Create README**

```markdown
# CSV Automation Tool

Automated Computer System Validation (CSV) test execution with AI-powered navigation.

## Features

- Execute test scenarios defined in Excel workbooks
- AI-powered navigation using Claude for intelligent UI interaction
- Full-screen screenshots at every step
- Word document evidence generation for audit compliance
- Navigation caching for cost optimization
- GitHub Actions integration for CI/CD

## Installation

```bash
pip install -e .
playwright install chromium
```

## Usage

### Command Line

```bash
# Run tests against an application
csvtool run \
  --url "https://app.example.com" \
  --workbook "scenarios/test-scenarios.xlsx"

# With options
csvtool run \
  --url "https://app.example.com" \
  --workbook "scenarios/test-scenarios.xlsx" \
  --output "custom_output/" \
  --no-headless  # Show browser window
```

### GitHub Actions

1. Add your Anthropic API key as a repository secret: `ANTHROPIC_API_KEY`
2. Go to Actions > Run CSV Tests
3. Enter the application URL and select the workbook
4. Click "Run workflow"
5. Download the evidence document from workflow artifacts

## Workbook Format

### Sheet: Credentials

| Role | Username | Password |
|------|----------|----------|
| Admin | admin@example.com | password |
| Quality Specialist | qs@example.com | password |

### Sheet: Test Scenarios

| Test ID | Role | Title | Test Instructions | Expected Result |
|---------|------|-------|-------------------|-----------------|
| TEST-001 | Admin | Verify settings access | 1. Navigate to settings\n2. Verify page loads | Settings page displayed |

## Environment Variables

- `ANTHROPIC_API_KEY` - Required. Your Anthropic API key for Claude.

## Development

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Run linting
ruff check src/
```

## License

MIT
```

- [ ] **Step 3: Commit**

```bash
git add .github/workflows/run-tests.yml README.md
git commit -m "feat: add GitHub Actions workflow and documentation"
```

---

## Chunk 10: Integration Testing

### Task 10.1: Add Integration Tests

**Files:**
- Create: `tests/test_integration.py`

- [ ] **Step 1: Write integration test**

```python
# tests/test_integration.py
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
            mock_browser.screenshot = AsyncMock(
                return_value=config.output_dir / "screenshots" / "test.png"
            )
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

            # Create screenshot directory and dummy file
            screenshot_dir = config.output_dir / "screenshots"
            screenshot_dir.mkdir(parents=True, exist_ok=True)
            dummy_screenshot = screenshot_dir / "test.png"

            # Create minimal PNG
            from PIL import Image
            img = Image.new("RGB", (100, 100), color="white")
            img.save(dummy_screenshot)

            mock_browser.screenshot = AsyncMock(return_value=dummy_screenshot)

            # Run executor
            executor = Executor(config)
            summary = await executor.run()

            # Verify results
            assert summary.total_tests == 1
            assert summary.passed_tests == 1
            assert summary.failed_tests == 0

            # Generate evidence
            generator = EvidenceGenerator(config.output_dir)
            doc_path = generator.generate(summary)

            assert doc_path.exists()
            assert doc_path.suffix == ".docx"

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
```

- [ ] **Step 2: Run integration tests**

```bash
pytest tests/test_integration.py -v
```

Expected: All PASS

- [ ] **Step 3: Run full test suite**

```bash
pytest --cov=csvtool --cov-report=term-missing
```

- [ ] **Step 4: Commit**

```bash
git add tests/test_integration.py
git commit -m "test: add integration tests for full workflow"
```

---

## Final Steps

### Task 11.1: Final Verification and Cleanup

- [ ] **Step 1: Run full test suite**

```bash
pytest -v --tb=short
```

Expected: All tests pass

- [ ] **Step 2: Verify CLI works**

```bash
csvtool --help
csvtool run --help
```

- [ ] **Step 3: Create sample workbook for testing**

```bash
mkdir -p scenarios
python tests/fixtures/create_sample_workbook.py
mv tests/fixtures/sample_workbook.xlsx scenarios/
```

- [ ] **Step 4: Final commit**

```bash
git add .
git commit -m "chore: finalize project structure"
```

- [ ] **Step 5: Tag release**

```bash
git tag -a v0.1.0 -m "Initial release"
```

---

## Summary

This plan implements the CSV Automation Tool in 10 chunks:

1. **Project Setup** - Dependencies, project structure, core models
2. **Excel Parser** - Workbook loading for credentials and scenarios
3. **Browser Automation** - Playwright wrapper for browser control
4. **AI Navigator** - Claude integration for intelligent navigation
5. **Navigation Cache** - Caching learned selectors
6. **Test Executor** - Orchestration of test execution
7. **Evidence Generator** - Word document generation
8. **CLI Interface** - Command-line interface with Click
9. **GitHub Actions** - CI/CD workflow configuration
10. **Integration Tests** - Full workflow testing

Each task follows TDD: write failing test → implement → verify → commit.
