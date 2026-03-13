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
