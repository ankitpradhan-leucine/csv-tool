# CSV Automation Tool - Design Specification

## Overview

A tool to automate Computer System Validation (CSV) test execution across multiple web applications. The tool executes test scenarios defined in Excel workbooks, captures screenshots as evidence at every step, and generates Word documents for audit compliance.

## Problem Statement

Currently, CSV testing requires:
- Manual execution of test scenarios by humans
- Human navigators who know each product's UI
- Manual screenshot capture during execution
- Manual compilation of evidence documents
- Separate VM deployments per product

This process is time-consuming, error-prone, and expensive.

## Solution

An automated tool that:
- Reads test scenarios from Excel workbooks
- Uses AI (Claude) to intelligently navigate different web UIs
- Captures full-screen screenshots at every action
- Generates Word documents with full traceability
- Runs via GitHub Actions with minimal inputs

## Inputs

The tool requires exactly two inputs:

| Input | Description |
|-------|-------------|
| **URL** | The application URL to test |
| **Workbook** | Excel file containing credentials and test scenarios |

## Workbook Structure

Each product has its own Excel workbook with the following sheets:

### Sheet 1: Credentials

| Role | Username | Password |
|------|----------|----------|
| Admin | admin@company.com | ••••• |
| Quality Specialist | qs@company.com | ••••• |
| Operator | operator@company.com | ••••• |
| Viewer | viewer@company.com | ••••• |

### Sheet 2+: Test Scenarios

| Test ID | Role | Title | Test Instructions | Expected Result |
|---------|------|-------|-------------------|-----------------|
| URAP-01 | Quality Specialist | Creates protocol with empty field | 1. Open protocol creation form<br>2. Enter details and leave mandatory field blank<br>3. Submit | Error shown |

Test instructions are written in natural language as numbered steps within a single cell.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                     GitHub Actions (Trigger)                     │
│   - Manual dispatch with URL + Workbook inputs                   │
└─────────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                      CSV Automation Tool                         │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────┐  │
│  │ Excel Parser│  │  Executor   │  │   Evidence Generator    │  │
│  │             │  │             │  │                         │  │
│  │ - Credentials│ │ - Playwright│  │ - Word doc generation   │  │
│  │ - Scenarios │  │ - Claude AI │  │ - Screenshot embedding  │  │
│  │             │  │ - Caching   │  │ - Traceability mapping  │  │
│  └─────────────┘  └─────────────┘  └─────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
                                │
                                ▼
                    ┌──────────────────────┐
                    │   Web Application    │
                    │   (via provided URL) │
                    └──────────────────────┘
```

### Components

1. **Excel Parser** - Loads workbook, parses credentials and test scenarios
2. **Executor** - Orchestrates test execution using Playwright + Claude
3. **Browser (Playwright)** - Headless browser control, screenshots, actions
4. **AI Navigator (Claude)** - Interprets instructions, identifies elements, generates test data
5. **Navigation Cache** - Stores learned selectors to reduce API calls
6. **Evidence Generator** - Creates Word documents with embedded screenshots

## Execution Flow

```
1. INITIALIZATION
   ├── Load workbook
   ├── Parse credentials sheet
   ├── Parse test scenarios
   └── Launch headless browser (Playwright)

2. FOR EACH TEST SCENARIO:
   │
   ├── 2a. LOGIN
   │   ├── Navigate to URL
   │   ├── Get credentials for required Role
   │   ├── Claude identifies login form → enters credentials
   │   └── Screenshot: "Login successful"
   │
   ├── 2b. EXECUTE STEPS (for each instruction line)
   │   ├── Check navigation cache
   │   ├── If cache HIT → use cached selector
   │   ├── If cache MISS → Claude analyzes screen → returns action
   │   ├── Execute action (click/type/etc)
   │   ├── Capture full-screen screenshot
   │   ├── Update cache with learned selector
   │   └── Wait for UI to stabilize
   │
   ├── 2c. VERIFY EXPECTED RESULT
   │   ├── Claude analyzes final screenshot
   │   ├── Compares against expected result
   │   └── Returns PASS/FAIL with reasoning
   │
   ├── 2d. LOGOUT / RESET
   │   └── Clear session for next test
   │
   └── 2e. ON FAILURE → Stop execution, save evidence collected

3. GENERATE EVIDENCE
   ├── Create Word document
   ├── For each test: ID, title, steps, screenshots, result
   └── Save to output folder
```

## Navigation Cache

The cache reduces Claude API calls by storing successful navigation paths.

### Cache Structure (JSON)

```json
{
  "cache_version": "1.0",
  "product_url": "https://app.company.com",
  "last_updated": "2026-03-13T14:30:00Z",
  "navigations": {
    "open protocol creation form": {
      "selector": "[data-testid='create-protocol-btn']",
      "action": "click",
      "success_count": 23,
      "last_success": "2026-03-13T14:30:05Z"
    }
  }
}
```

### Cache Behavior

| Scenario | Behavior |
|----------|----------|
| Instruction not in cache | Claude analyzes → executes → saves to cache |
| Cache hit, selector works | Use cached selector (no API call) |
| Cache hit, selector fails | Fall back to Claude → update cache |

### Cost Impact

- First run: ~$80-100 per product (full AI usage)
- Subsequent runs: ~$20-50 (mostly cache hits)
- Self-healing: UI changes trigger automatic cache updates

## Evidence Document Format

Generated Word document (.docx) structure:

```
┌─────────────────────────────────────────────────────────────────┐
│                    CSV TEST EVIDENCE REPORT                      │
│                                                                  │
│  Application URL: https://app.company.com                        │
│  Workbook: product-a.xlsx                                        │
│  Execution Date: 2026-03-13 14:30:00                            │
│  Total Tests: 45 | Passed: 44 | Failed: 1                       │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│  TEST: URAP-TEST-01                                              │
│  Title: Quality Specialist creates protocol with empty field     │
│  Role: Quality Specialist                                        │
│  Status: PASSED                                                  │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│  Step 1: Open protocol creation form                             │
│  [SCREENSHOT]                                                    │
│  Timestamp: 14:30:05                                            │
│                                                                  │
│  Step 2: Enter details and leave mandatory field blank           │
│  [SCREENSHOT]                                                    │
│  Timestamp: 14:30:12                                            │
│                                                                  │
│  Expected Result: Error message shown                            │
│  Observed Result: Error displayed - "Type field is required"     │
│  Verification: PASSED                                            │
└─────────────────────────────────────────────────────────────────┘
```

File naming: `CSV-Evidence-{workbook-name}-{date}-{time}.docx`

## Error Handling

| Error Type | Behavior | Evidence Captured |
|------------|----------|-------------------|
| Element not found | Retry with Claude → stop if still fails | Last screenshot + error |
| Login failed | Stop immediately | Login page screenshot |
| Page timeout | Retry once → stop | Current state screenshot |
| Unexpected popup | Claude dismisses if possible → continue | Popup screenshot |
| Network error | Retry 3x with backoff → stop | Error logged |
| Claude API error | Retry 3x → stop | Last screenshot |
| Expected result mismatch | Mark FAILED → stop | Final screenshot + comparison |

On any failure, execution stops and partial evidence is saved.

## GitHub Actions Integration

### Workflow Configuration

```yaml
name: Run CSV Tests

on:
  workflow_dispatch:
    inputs:
      application_url:
        description: 'Application URL'
        required: true
        type: string

      workbook:
        description: 'Test workbook'
        required: true
        type: choice
        options:
          - scenarios/workbook-1.xlsx
          - scenarios/workbook-2.xlsx
          - scenarios/workbook-3.xlsx

jobs:
  run-csv-tests:
    runs-on: self-hosted

    steps:
      - uses: actions/checkout@v4

      - name: Run CSV Automation
        env:
          ANTHROPIC_API_KEY: ${{ secrets.ANTHROPIC_API_KEY }}
        run: |
          python -m csvtool run \
            --url "${{ inputs.application_url }}" \
            --workbook "${{ inputs.workbook }}"

      - name: Upload Evidence
        uses: actions/upload-artifact@v4
        with:
          name: csv-evidence-${{ github.run_number }}
          path: output/*.docx
```

### User Experience

User triggers workflow from GitHub UI:
1. Select "Run workflow"
2. Enter Application URL
3. Select Workbook from dropdown
4. Click "Run workflow"
5. Download evidence document from artifacts when complete

## Technology Stack

| Component | Technology | Rationale |
|-----------|------------|-----------|
| Language | Python 3.11+ | Library support, maintainability |
| Browser Automation | Playwright | Fast, reliable, headless, screenshots |
| AI Integration | Anthropic Claude API | Vision + reasoning capabilities |
| Excel Parsing | openpyxl | Native .xlsx support |
| Word Generation | python-docx | Formatted .docx with images |
| CLI Interface | Click | Clean argument handling |
| Caching | JSON files | Simple, git-friendly |

### Project Structure

```
csv-automation/
├── src/
│   ├── __init__.py
│   ├── main.py              # CLI entry point
│   ├── excel_parser.py      # Workbook loading
│   ├── executor.py          # Test execution orchestration
│   ├── browser.py           # Playwright wrapper
│   ├── ai_navigator.py      # Claude integration
│   ├── cache.py             # Navigation cache management
│   ├── evidence.py          # Word document generation
│   └── models.py            # Data classes
├── scenarios/
│   └── *.xlsx               # Test workbooks
├── cache/
│   └── *.json               # Navigation caches
├── output/
│   └── *.docx               # Evidence documents
├── .github/
│   └── workflows/
│       └── run-tests.yml
├── requirements.txt
├── pyproject.toml
└── README.md
```

### Dependencies

```
playwright>=1.40.0
anthropic>=0.18.0
openpyxl>=3.1.0
python-docx>=1.1.0
click>=8.1.0
pydantic>=2.0.0
```

## Cost Estimates

### Scale

- 5 products
- 1000 test scenarios per product
- Monthly execution

### Monthly Costs

| Item | First Month | Subsequent Months |
|------|-------------|-------------------|
| Claude API | ~$400-500 | ~$100-200 (cached) |
| Self-hosted runner | Free | Free |
| VM (if dedicated) | ~$30-50 | ~$30-50 |
| **Total** | ~$450-550 | ~$150-250 |

## Test Data Generation

When test instructions require entering data (e.g., "Enter details"), Claude generates contextually appropriate dummy data:

- Identifies field types from labels/placeholders
- Generates realistic values (names, dates, emails, numbers)
- For "leave field blank" instructions, identifies mandatory fields and skips one

## Security Considerations

1. **Credentials** - Stored in Excel workbooks; repos should be private
2. **API Keys** - Stored as GitHub secrets, never in code
3. **Evidence Documents** - May contain sensitive data; handle per organization policy
4. **Network Access** - Self-hosted runner needs access to application URLs

## Success Criteria

1. Tool executes test scenarios without human intervention
2. Screenshots captured at every step showing actual URLs
3. Evidence documents meet audit compliance requirements
4. 80%+ cache hit rate after initial learning run
5. Monthly cost under $300 after optimization

## Future Enhancements (Out of Scope)

- Parallel test execution
- Multiple browser support
- Video recording option
- Integration with test management tools
- Automatic Excel result population
