# Complete Usage Guide - CSV Automation Tool

This guide covers everything from initial setup to running tests every time.

---

## 📦 PART 1: ONE-TIME SETUP (Do this once)

### Step 1: Install System Prerequisites

```bash
# Install pip and venv (requires sudo)
sudo apt-get update
sudo apt-get install python3-pip python3.12-venv
```

**Verify installation:**
```bash
python3 --version   # Should show Python 3.12.3 or higher
python3 -m pip --version   # Should show pip version
```

---

### Step 2: Navigate to Project Directory

```bash
cd /home/ankit/csvtool/.worktrees/csv-automation
```

**Verify you're in the right place:**
```bash
pwd
# Should show: /home/ankit/csvtool/.worktrees/csv-automation

ls
# Should show: src/, tests/, pyproject.toml, README.md, etc.
```

---

### Step 3: Create Virtual Environment

```bash
python3 -m venv venv
```

**What this does:** Creates an isolated Python environment in the `venv/` folder.

---

### Step 4: Activate Virtual Environment

```bash
source venv/bin/activate
```

**You'll see:** Your terminal prompt changes to show `(venv)` at the beginning.

**Example:**
```
Before: ankit@machine:~/csvtool/.worktrees/csv-automation$
After:  (venv) ankit@machine:~/csvtool/.worktrees/csv-automation$
```

---

### Step 5: Install the Tool and Dependencies

```bash
pip install -e .
```

**This installs:**
- csvtool (the tool itself)
- playwright (browser automation)
- anthropic (Claude AI)
- openpyxl (Excel parsing)
- python-docx (Word document generation)
- click (CLI framework)
- pydantic (data validation)

**Time:** Takes 1-2 minutes. You'll see download progress bars.

---

### Step 6: Install Chromium Browser

```bash
playwright install chromium
```

**This downloads:** Chrome browser for testing (~280 MB).

**Time:** Takes 2-3 minutes depending on internet speed.

---

### Step 7: Verify Installation

```bash
csvtool --version
```

**Expected output:**
```
CSV Automation Tool, version 0.1.0
```

**If this shows correctly, setup is complete!** ✅

---

### Step 8: Create Sample Test Workbook (Optional)

```bash
python3 create_sample_workbook.py
```

**This creates:** `scenarios/sample-test.xlsx` with example tests.

---

## 🎯 PART 2: RUNNING TESTS (Do this every time)

### Step 1: Open Terminal and Navigate to Project

```bash
cd /home/ankit/csvtool/.worktrees/csv-automation
```

---

### Step 2: Activate Virtual Environment

```bash
source venv/bin/activate
```

**Important:** You must do this every time you open a new terminal!

**Check:** Your prompt should show `(venv)` at the start.

---

### Step 3: Set Your Anthropic API Key

```bash
export ANTHROPIC_API_KEY="sk-ant-your-actual-key-here"
```

**Replace** `sk-ant-your-actual-key-here` with your real API key from https://console.anthropic.com/

**Note:** This is temporary - only lasts for this terminal session.

**To make it permanent** (optional):
```bash
echo 'export ANTHROPIC_API_KEY="sk-ant-your-actual-key-here"' >> ~/.bashrc
source ~/.bashrc
```

---

### Step 4: Run a Test

**Option A: Test with example.com (visible browser)**
```bash
csvtool run \
  --url "https://example.com" \
  --workbook "scenarios/sample-test.xlsx" \
  --no-headless
```

**What happens:**
- Browser window opens (you can watch it)
- Navigates to example.com
- Executes test scenarios
- Takes screenshots
- Generates Word document

**Option B: Test with your real application (headless mode)**
```bash
csvtool run \
  --url "https://your-actual-app.com" \
  --workbook "scenarios/your-test-scenarios.xlsx" \
  --headless
```

**Option C: Test with custom output directory**
```bash
csvtool run \
  --url "https://example.com" \
  --workbook "scenarios/sample-test.xlsx" \
  --output "custom-results" \
  --cache-dir "custom-cache"
```

---

### Step 5: Check Results

**View evidence document:**
```bash
ls -lh output/*.docx
```

**View screenshots:**
```bash
ls output/screenshots/
```

**View cache (learned selectors):**
```bash
cat cache/*.json
```

**Open evidence document:**
```bash
# On WSL with Windows
explorer.exe output/

# Or copy to Windows
cp output/*.docx /mnt/c/Users/YourName/Desktop/
```

---

## 📝 PART 3: CREATING YOUR OWN TEST WORKBOOK

### Excel Workbook Structure

Your Excel file needs at least 2 sheets:

#### Sheet 1: "Credentials"

| Role | Username | Password |
|------|----------|----------|
| Admin | admin@yourapp.com | password123 |
| Quality Specialist | qs@yourapp.com | qspass456 |
| Operator | operator@yourapp.com | oppass789 |

**Rules:**
- First sheet must be named "Credentials"
- Must have exactly 3 columns: Role, Username, Password
- Add as many users as you need

---

#### Sheet 2+: Test Scenarios (any name)

| Test ID | Role | Title | Test Instructions | Expected Result |
|---------|------|-------|-------------------|-----------------|
| TEST-001 | Admin | Verify login | 1. Navigate to login page<br>2. Enter credentials<br>3. Click login button | Dashboard is displayed |
| TEST-002 | Quality Specialist | Create record | 1. Click New Record<br>2. Fill all fields<br>3. Click Save | Record saved successfully |

**Rules:**
- Can have multiple sheets with test scenarios
- Each sheet needs 5 columns: Test ID, Role, Title, Test Instructions, Expected Result
- Test Instructions: Write numbered steps (1. Do this\n2. Do that\n3. Do another)
- Role must match a role from the Credentials sheet

---

### Save Your Workbook

```bash
# Save in scenarios/ folder
scenarios/my-app-tests.xlsx
```

---

## 🔧 PART 4: COMMON COMMANDS REFERENCE

### Check Tool Version
```bash
csvtool --version
```

### View Help
```bash
csvtool --help
csvtool run --help
```

### Run with All Options
```bash
csvtool run \
  --url "https://app.example.com" \
  --workbook "scenarios/test.xlsx" \
  --output "results/2024-03-15" \
  --cache-dir "cache" \
  --no-headless \
  --continue-on-failure \
  --api-key "sk-ant-your-key"
```

**Options explained:**
- `--url`: Application URL to test (required)
- `--workbook`: Path to Excel file (required)
- `--output`: Where to save results (default: output/)
- `--cache-dir`: Where to save cache (default: cache/)
- `--headless`: Run browser in background (default)
- `--no-headless`: Show browser window
- `--stop-on-failure`: Stop at first failure (default)
- `--continue-on-failure`: Run all tests even if some fail
- `--api-key`: Anthropic API key (or use environment variable)

### View Cached Selectors
```bash
cat cache/*.json | python3 -m json.tool
```

### Clean Up Old Results
```bash
rm -rf output/
rm -rf cache/
```

---

## 🚨 PART 5: TROUBLESHOOTING

### Error: "ANTHROPIC_API_KEY must be provided"

**Problem:** API key not set

**Solution:**
```bash
export ANTHROPIC_API_KEY="sk-ant-your-key"
```

---

### Error: "csvtool: command not found"

**Problem:** Virtual environment not activated

**Solution:**
```bash
source venv/bin/activate
```

---

### Error: "ModuleNotFoundError: No module named 'playwright'"

**Problem:** Dependencies not installed

**Solution:**
```bash
pip install -e .
```

---

### Error: "Executable doesn't exist at /home/.../chromium"

**Problem:** Playwright browser not installed

**Solution:**
```bash
playwright install chromium
```

---

### Error: "Workbook not found"

**Problem:** Wrong path to Excel file

**Solution:**
```bash
# Check file exists
ls scenarios/

# Use correct path
csvtool run --url "..." --workbook "scenarios/correct-name.xlsx"
```

---

### Browser doesn't open (with --no-headless)

**Problem:** Running on WSL without X server

**Solutions:**
1. Install X server (VcXsrv or Xming on Windows)
2. Use headless mode: `--headless` (default)
3. Run on a machine with display

---

### Tests are slow / expensive

**Problem:** Not using cache effectively

**Solution:**
- Run tests multiple times - second run uses cache (cheaper)
- Keep cache/ folder - don't delete between runs
- Check cache hit rate: `cat cache/*.json`

---

## 📊 PART 6: UNDERSTANDING THE OUTPUT

### Evidence Document Structure

The Word document contains:

1. **Header**
   - Application URL
   - Workbook name
   - Execution date/time

2. **Summary Table**
   - Total tests
   - Passed/Failed
   - Success rate

3. **Individual Test Results** (for each test)
   - Test ID and Title
   - Role used
   - Status (PASSED/FAILED)
   - Expected result
   - Each step with:
     - Instruction
     - Timestamp
     - Screenshot (embedded)
   - Verification result

---

### Cache File Format

```json
{
  "cache_version": "1.0",
  "product_url": "https://app.example.com",
  "last_updated": "2026-03-15T14:30:00Z",
  "navigations": {
    "click login button": {
      "selector": "button[type='submit']",
      "action": "click",
      "success_count": 15,
      "last_success": "2026-03-15T14:30:05Z"
    }
  }
}
```

**What it means:**
- Each instruction is cached with its selector
- `success_count`: How many times it worked
- Higher count = more reliable selector

---

## 🎓 PART 7: BEST PRACTICES

### Writing Good Test Instructions

✅ **Good:**
```
1. Click the 'New User' button
2. Enter 'John Doe' in the name field
3. Enter 'john@example.com' in email field
4. Click 'Save' button
```

❌ **Bad:**
```
Create a new user
```

**Why:** AI needs specific, step-by-step instructions.

---

### Using Roles Effectively

**Tip:** Use different roles to test different permissions.

```
Admin role:
- TEST-001: Create user (should work)
- TEST-002: Delete user (should work)

Viewer role:
- TEST-003: View users (should work)
- TEST-004: Delete user (should fail with error)
```

---

### Expected Results

✅ **Good:**
```
User created successfully, confirmation message displayed
```

✅ **Good:**
```
Error message: "Permission denied"
```

❌ **Bad:**
```
Success
```

**Why:** Specific results help AI verify correctly.

---

## 📈 PART 8: COST ESTIMATION

### First Run (Learning Phase)
- **Per product:** ~$80-100
- AI analyzes every element (no cache)

### Subsequent Runs (Cache Hits)
- **Per product:** ~$20-50
- Uses cached selectors (60-80% savings)

### Monthly Cost (5 products, monthly execution)
- **First month:** ~$400-500
- **After caching:** ~$100-250/month

**Tips to reduce costs:**
1. Keep cache/ folder between runs
2. Run tests in batches
3. Use --stop-on-failure to catch issues early
4. Reuse cache for similar applications

---

## 🔄 PART 9: QUICK REFERENCE COMMANDS

### Daily Workflow

```bash
# 1. Navigate
cd /home/ankit/csvtool/.worktrees/csv-automation

# 2. Activate
source venv/bin/activate

# 3. Set key (if not permanent)
export ANTHROPIC_API_KEY="sk-ant-your-key"

# 4. Run test
csvtool run --url "https://yourapp.com" --workbook "scenarios/tests.xlsx"

# 5. Check results
ls -lh output/*.docx
```

### Copy-Paste Template

```bash
cd /home/ankit/csvtool/.worktrees/csv-automation && \
source venv/bin/activate && \
export ANTHROPIC_API_KEY="YOUR-KEY-HERE" && \
csvtool run --url "YOUR-URL" --workbook "YOUR-WORKBOOK.xlsx"
```

**Just replace:**
- `YOUR-KEY-HERE` with your API key
- `YOUR-URL` with your application URL
- `YOUR-WORKBOOK.xlsx` with your test file

---

## 📤 PART 10: SHARING WITH OTHERS

### Share the Tool

**Option 1: Share this directory**
```bash
# Zip the project (without venv to save space)
cd /home/ankit/csvtool/.worktrees/
tar -czf csv-automation-tool.tar.gz csv-automation/ --exclude=csv-automation/venv

# Share csv-automation-tool.tar.gz with others
```

**Option 2: Share via Git**
```bash
cd /home/ankit/csvtool/.worktrees/csv-automation
git push -u origin feature/csv-automation-implementation

# Others can clone and follow QUICKSTART.md
```

---

### What to Share

1. **This guide** (COMPLETE-USAGE-GUIDE.md)
2. **Quick start** (QUICKSTART.md)
3. **Sample workbook** (scenarios/sample-test.xlsx)
4. **The tool** (entire project directory)

**What NOT to share:**
- Your API key
- Your test results (may contain sensitive data)
- Your actual test workbooks (may contain credentials)

---

## 🎉 You're All Set!

**Next steps:**
1. Get your Anthropic API key from https://console.anthropic.com/
2. Run the commands in PART 2
3. Check the evidence document
4. Create your own test workbook
5. Test your actual application

**Questions?**
- Check troubleshooting (PART 5)
- Review examples (PART 7)
- See SETUP-GUIDE.md for more details

**Happy Testing!** 🚀
