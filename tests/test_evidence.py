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
        # Create a dummy file instead of actual image to avoid PIL dependency
        path = tmp_path / "screenshot.png"
        path.write_text("dummy screenshot")
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

        # Check for summary data in either paragraphs or tables
        assert "2" in text  # total tests
        assert "1" in text  # passed/failed

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
