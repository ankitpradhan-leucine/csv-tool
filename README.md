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
| TEST-001 | Admin | Verify settings access | 1. Navigate to settings<br>2. Verify page loads | Settings page displayed |

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
