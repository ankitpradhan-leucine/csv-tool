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

    def test_workbook_name_stored(self, sample_workbook_path: Path):
        """Test that workbook name is stored in WorkbookData."""
        parser = ExcelParser(sample_workbook_path)
        data = parser.parse()
        assert data.workbook_name == "sample_workbook.xlsx"
