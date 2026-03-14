# Quick Start Guide

Get the CSV Automation Tool running in 5 minutes!

## Prerequisites Check

```bash
# Check Python version (need 3.11+)
python3 --version

# Check if pip is available
python3 -m pip --version
```

**If pip is not available:** See SETUP-GUIDE.md for installation instructions.

## Installation (One-Time Setup)

```bash
# 1. Navigate to project
cd /home/ankit/csvtool/.worktrees/csv-automation

# 2. Install the tool and dependencies
pip install -e .

# 3. Install browser
playwright install chromium

# 4. Set your API key (get from https://console.anthropic.com/)
export ANTHROPIC_API_KEY="your-key-here"

# 5. Create sample test workbook
python3 create_sample_workbook.py
```

## Run Your First Test

```bash
# Run with visible browser (so you can see what's happening)
csvtool run \
  --url "https://example.com" \
  --workbook "scenarios/sample-test.xlsx" \
  --no-headless
```

**What happens:**
1. Browser opens and navigates to example.com
2. Logs in (attempts to - example.com has no login)
3. Executes each test step
4. Takes screenshots at every step
5. Generates Word document with evidence

## Check Results

```bash
# View evidence document
ls -lh output/*.docx

# View screenshots
ls output/screenshots/

# View cache (learned selectors)
cat cache/*.json
```

Open the `.docx` file to see the complete audit trail!

## Next Steps

1. **Test with your application:**
   - Edit `scenarios/sample-test.xlsx`
   - Update URL to your application
   - Add real credentials
   - Write your test scenarios

2. **Run in headless mode** (background):
   ```bash
   csvtool run --url "https://your-app.com" --workbook "scenarios/sample-test.xlsx"
   ```

3. **Deploy to GitHub Actions:**
   - Push branch to GitHub
   - Add `ANTHROPIC_API_KEY` as repository secret
   - Run workflow from Actions tab

## Troubleshooting

**"ANTHROPIC_API_KEY must be provided"**
```bash
export ANTHROPIC_API_KEY="your-key"
```

**"ModuleNotFoundError"**
```bash
pip install -e .
```

**"Executable doesn't exist"**
```bash
playwright install chromium
```

Need more help? See **SETUP-GUIDE.md** for detailed instructions.
