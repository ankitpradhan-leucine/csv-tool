"""Tests for data models."""

import pytest
from datetime import datetime
from pathlib import Path

from csvtool.models import (
    Credential,
    TestStep,
    TestScenario,
    StepResult,
    TestResult,
    ExecutionSummary,
)


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


class TestExecutionSummary:
    """Tests for ExecutionSummary model."""

    def test_execution_summary_creation(self):
        """Test creating an execution summary."""
        summary = ExecutionSummary(
            url="http://example.com",
            workbook_name="test_scenarios.xlsx",
            execution_date=datetime(2026, 3, 13, 15, 0, 0),
            total_tests=10,
            passed_tests=8,
            failed_tests=2,
            test_results=[]
        )
        assert summary.url == "http://example.com"
        assert summary.workbook_name == "test_scenarios.xlsx"
        assert summary.total_tests == 10

    def test_success_rate_calculation(self):
        """Test success rate property calculation."""
        summary = ExecutionSummary(
            url="http://example.com",
            workbook_name="test.xlsx",
            execution_date=datetime(2026, 3, 13, 15, 0, 0),
            total_tests=10,
            passed_tests=8,
            failed_tests=2,
            test_results=[]
        )
        assert summary.success_rate == 80.0

    def test_success_rate_all_passed(self):
        """Test success rate when all tests pass."""
        summary = ExecutionSummary(
            url="http://example.com",
            workbook_name="test.xlsx",
            execution_date=datetime(2026, 3, 13, 15, 0, 0),
            total_tests=5,
            passed_tests=5,
            failed_tests=0,
            test_results=[]
        )
        assert summary.success_rate == 100.0

    def test_success_rate_all_failed(self):
        """Test success rate when all tests fail."""
        summary = ExecutionSummary(
            url="http://example.com",
            workbook_name="test.xlsx",
            execution_date=datetime(2026, 3, 13, 15, 0, 0),
            total_tests=5,
            passed_tests=0,
            failed_tests=5,
            test_results=[]
        )
        assert summary.success_rate == 0.0

    def test_success_rate_zero_tests(self):
        """Test success rate with no tests."""
        summary = ExecutionSummary(
            url="http://example.com",
            workbook_name="test.xlsx",
            execution_date=datetime(2026, 3, 13, 15, 0, 0),
            total_tests=0,
            passed_tests=0,
            failed_tests=0,
            test_results=[]
        )
        assert summary.success_rate == 0.0
