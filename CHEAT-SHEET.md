# CSV Automation Tool - Cheat Sheet

Quick reference for running tests.

---

## 🚀 First Time Setup (Once)

```bash
cd /home/ankit/csvtool/.worktrees/csv-automation
python3 -m venv venv
source venv/bin/activate
pip install -e .
playwright install chromium
```

---

## ▶️ Run Test (Every Time)

```bash
cd /home/ankit/csvtool/.worktrees/csv-automation
source venv/bin/activate
export ANTHROPIC_API_KEY="your-key-here"
csvtool run --url "https://example.com" --workbook "scenarios/sample-test.xlsx"
```

---

## 📋 Common Commands

```bash
# Show version
csvtool --version

# Show help
csvtool run --help

# Run with visible browser
csvtool run --url "..." --workbook "..." --no-headless

# Run and continue on failure
csvtool run --url "..." --workbook "..." --continue-on-failure

# Check results
ls -lh output/*.docx
ls output/screenshots/

# View cache
cat cache/*.json
```

---

## 🔧 Troubleshooting

| Error | Fix |
|-------|-----|
| `command not found: csvtool` | `source venv/bin/activate` |
| `ANTHROPIC_API_KEY must be provided` | `export ANTHROPIC_API_KEY="..."` |
| `ModuleNotFoundError` | `pip install -e .` |
| `Executable doesn't exist` | `playwright install chromium` |

---

## 📝 Workbook Format

**Sheet 1: Credentials**
```
Role | Username | Password
Admin | admin@app.com | pass123
```

**Sheet 2+: Tests**
```
Test ID | Role | Title | Test Instructions | Expected Result
TEST-001 | Admin | Login | 1. Go to login\n2. Enter creds\n3. Click login | Dashboard shown
```

---

## 💡 Tips

- Keep `cache/` folder for cost savings
- Use `--no-headless` to debug
- First run expensive (~$80), subsequent cheaper (~$20)
- Write specific instructions (not vague)

---

**Full Guide:** See COMPLETE-USAGE-GUIDE.md
