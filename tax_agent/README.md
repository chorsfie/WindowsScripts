# 🏦 UK Tax Analysis Agent — Local & Private

A local AI agent that reads your bank statement PDFs and analyses UK tax implications.
**100% private — everything runs on your machine. No data ever leaves your computer.**

---

## What it analyses

- **Income Tax** — salary, freelance, rental income
- **Savings Interest** — vs your Personal Savings Allowance (£1,000 / £500)
- **Dividends** — vs £500 dividend allowance, ISA vs non-ISA
- **Capital Gains** — vs £3,000 annual exempt amount
- **ISA Usage** — deposits vs £20,000 annual allowance
- **Pension Contributions** — tax relief estimates
- **Self Assessment triggers** — whether you need to file
- **Tax efficiency suggestions** — specific to your situation

---

## Quick Setup

### Step 1 — Install Ollama (local LLM)

**Mac / Windows:**
→ Download from https://ollama.com and run the installer

**Linux:**
```bash
curl -fsSL https://ollama.ai/install.sh | sh
```

### Step 2 — Download a model

Choose one based on your RAM:

| Model | RAM needed | Quality | Command |
|-------|-----------|---------|---------|
| `phi4-mini` | ~3 GB | Good | `ollama pull phi4-mini` |
| `llama3.1` | ~5 GB | Very good | `ollama pull llama3.1` |
| `llama3.3` | ~40 GB | Excellent | `ollama pull llama3.3` |

**Recommended for most people:**
```bash
ollama pull llama3.1
```

### Step 3 — Start Ollama

```bash
ollama serve
```
Leave this terminal open.

### Step 4 — Install Python dependencies

```bash
pip install pdfplumber requests
```

### Step 5 — Add your bank statements

Create a folder called `statements` in the same directory as `tax_agent.py`:
```
tax_agent/
  tax_agent.py
  statements/
    barclays_jan2024.pdf
    barclays_feb2024.pdf
    lloyds_savings_2024.pdf
    trading212_2024.pdf
    ...
```

### Step 6 — Run the agent

```bash
python tax_agent.py
```

Or with options:
```bash
# Custom folder and model
python tax_agent.py ./my_statements --model phi4-mini

# Save report with a custom name
python tax_agent.py ./statements --output 2024_tax_report.md

# Export normalized statement rows to CSV
python tax_agent.py ./statements --transactions-csv normalized_transactions.csv

# Skip LLM extraction stage and build events from CSV rows deterministically
python tax_agent.py ./statements --skip-llm-extraction --output tax_report_fast.md
```

---

## What to include in your statements folder

For the most complete analysis, include PDFs from:

- ✅ Current accounts (salary payments visible)
- ✅ Savings accounts (interest payments)
- ✅ Investment / brokerage accounts (dividends, sale proceeds)
- ✅ ISA accounts (deposits, income — to verify allowance usage)
- ✅ Pension statements (if you make personal contributions)
- ✅ P60 / P11D if you have them as PDFs

---

## Output

The agent produces:
1. **Terminal output** — analysis printed immediately
2. **`normalized_transactions.csv`** — normalized rows extracted from PDF tables
3. **`tax_report.md`** — full report saved to disk, including:
   - Raw extracted financial data (JSON)
   - Full tax analysis with estimated liabilities
   - Self Assessment guidance
   - Next steps

---

## Privacy & Security

- ✅ Ollama runs entirely offline on your machine
- ✅ No data is sent to any external server
- ✅ Your PDFs never leave your computer
- ✅ The report is saved locally only

---

## Limitations & Disclaimer

- This is an **estimate only** based on what appears in your bank statements
- Bank statements don't always show the full picture (e.g. employer pension contributions, P11D benefits)
- UK tax law is complex — always verify with HMRC or a qualified accountant before filing
- Tax thresholds in the script are for **2024/25** — update `UK_TAX` dict for future years

---

## Troubleshooting

**"Cannot connect to Ollama"**
→ Run `ollama serve` in a separate terminal

**"No PDF files found"**
→ Make sure PDFs are in the `./statements` folder (or specify your folder path)

**Slow analysis**
→ Use a smaller model: `--model phi4-mini`

**Poor extraction quality**
→ Some bank statement PDFs are scanned images — try running them through a free OCR tool first (e.g. Adobe Acrobat Reader's OCR feature)

---

*Built for UK tax year 2024/25. Rates: Personal Allowance £12,570 | Basic rate to £50,270 | CGT allowance £3,000 | Dividend allowance £500 | ISA allowance £20,000*
