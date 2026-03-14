# CSV Automation Tool - Local Setup Guide

## Prerequisites

You need a Python environment with pip installed. This guide assumes you have:
- Python 3.11 or higher
- pip package manager
- Internet connection

## Step-by-Step Installation

### 1. Install pip (if not already installed)

```bash
# On Ubuntu/Debian
sudo apt-get update
sudo apt-get install python3-pip

# Verify installation
python3 -m pip --version
```

### 2. Navigate to the project directory

```bash
cd /home/ankit/csvtool/.worktrees/csv-automation
```

### 3. Install the CSV Automation Tool

```bash
# Install in development mode (editable)
pip install -e .

# This installs:
# - playwright (browser automation)
# - anthropic (Claude AI)
# - openpyxl (Excel parsing)
# - python-docx (Word document generation)
# - click (CLI framework)
# - pydantic (data validation)
```

### 4. Install Playwright Browser

```bash
# Install Chromium browser for Playwright
playwright install chromium

# Install system dependencies for Playwright
playwright install-deps chromium
```

### 5. Set up your Anthropic API Key

You need a Claude API key from Anthropic:

```bash
# Option 1: Set as environment variable (temporary - for current session)
export ANTHROPIC_API_KEY="your-api-key-here"

# Option 2: Add to your ~/.bashrc or ~/.zshrc (permanent)
echo 'export ANTHROPIC_API_KEY="your-api-key-here"' >> ~/.bashrc
source ~/.bashrc

# Option 3: Pass directly to command (each time)
csvtool run --api-key "your-api-key-here" --url ... --workbook ...
```

**To get an API key:**
1. Go to https://console.anthropic.com/
2. Sign up or log in
3. Go to API Keys section
4. Create a new key

### 6. Verify Installation

```bash
# Check csvtool is installed
csvtool --version
# Should show: CSV Automation Tool, version 0.1.0

# Check help
csvtool --help
csvtool run --help
```

## Running Your First Test

### 1. Create a test workbook

Use the provided sample workbook:
```bash
# The sample workbook is at:
# /home/ankit/csvtool/.worktrees/csv-automation/scenarios/sample-test.xlsx
```

Or create your own Excel file with:
- Sheet 1: "Credentials" with columns: Role, Username, Password
- Sheet 2+: Test scenarios with columns: Test ID, Role, Title, Test Instructions, Expected Result

### 2. Run a simple test

```bash
# Test with example.com (visible browser)
csvtool run \
  --url "https://example.com" \
  --workbook "scenarios/sample-test.xlsx" \
  --no-headless
```

This will:
1. Open a browser window (--no-headless shows the browser)
2. Navigate to example.com
3. Execute the test scenarios
4. Take screenshots at each step
5. Generate a Word document with evidence

### 3. Check the output

```bash
# Evidence document
ls -lh output/*.docx

# Screenshots
ls -lh output/screenshots/

# Navigation cache
ls -lh cache/*.json
```

## Common Issues and Solutions

### "ANTHROPIC_API_KEY must be provided"

**Problem:** API key not set

**Solution:** Export the environment variable or pass --api-key flag

```bash
export ANTHROPIC_API_KEY="your-key"
```

### "playwright._impl._api_types.Error: Executable doesn't exist"

**Problem:** Playwright browser not installed

**Solution:** Run playwright install

```bash
playwright install chromium
playwright install-deps chromium
```

### "ModuleNotFoundError: No module named 'openpyxl'"

**Problem:** Dependencies not installed

**Solution:** Reinstall the package

```bash
pip install -e .
```

### "Permission denied" when installing

**Problem:** Need sudo access or virtual environment

**Solution:** Use a virtual environment

```bash
# Create virtual environment
python3 -m venv venv

# Activate it
source venv/bin/activate

# Install the tool
pip install -e .
```

## Testing with Real Applications

Once you've verified the tool works with example.com, test with your actual applications:

```bash
csvtool run \
  --url "https://your-app.company.com" \
  --workbook "scenarios/your-test-scenarios.xlsx" \
  --output "results/$(date +%Y%m%d)" \
  --headless  # Run in background
```

## Next Steps

After successful local testing:
1. Commit any changes to scenarios or configuration
2. Push the branch to GitHub: `git push -u origin feature/csv-automation-implementation`
3. Create a Pull Request
4. Set up GitHub Actions with your API key as a secret
5. Run automated tests via GitHub Actions workflow

## Support

If you encounter issues:
1. Check that all dependencies are installed: `pip list`
2. Verify Python version: `python3 --version` (need 3.11+)
3. Check Playwright: `playwright --version`
4. Review error logs in the evidence document
5. Check screenshots to see what the browser saw
