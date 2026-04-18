#!/usr/bin/env python3
"""
================================================================================
  UK TAX ANALYSIS AGENT
  Version: 1.3  |  Author: Built with Claude (Anthropic)  |  Licence: MIT
================================================================================

PURPOSE
-------
This script reads your bank statement PDFs and uses a locally-running AI model
to extract financial events (income, savings interest, dividends, capital gains,
ISA deposits, pension contributions) and then produces a UK tax analysis report.

Everything runs 100% on your own machine. No data is uploaded anywhere.

HOW IT WORKS  (the pipeline, step by step)
------------------------------------------
  1. DISCOVER
     The script walks your statements folder (and all subfolders) looking for
     PDF files. It deduplicates them (Google Drive can sometimes expose the same
     file twice via different paths) and filters to recent years only so the AI
     isn't overwhelmed by old irrelevant data.

  2. EXTRACT TEXT FROM PDFs
     Each PDF is opened with pdfplumber. The library tries two strategies:
       a) Table extraction  -- best for bank statements, which are laid out in
          columns (date | description | debit | credit | balance).
       b) Plain text        -- fallback for statements that don't use tables.
     The raw text from each page is joined into a single string per file.

  3. BATCH AND SEND TO LOCAL AI
     A local AI model (Meta Llama 3.1 8B, running inside LM Studio) has a
     limited "context window" -- it can only read ~4096 tokens (~3000 words) at
     once. Sending all 50+ statements at once would exceed that and cause a
     400 Bad Request error.

     Instead, the script groups statements into small batches (~3000 characters
     each) and sends them one batch at a time. The AI is asked to return
     structured JSON for each batch (income events, interest payments, etc.).

  4. MERGE BATCH RESULTS
     The JSON results from every batch are merged into a single master
     dictionary, so we end up with one complete picture of the year's finances
     even though it was processed in chunks.

  5. TAX ANALYSIS
     The merged financial data is sent (in one request) to the AI along with
     the current UK tax thresholds. The AI produces a plain-English analysis
     covering income tax, savings interest, dividends, CGT, ISA usage,
     pension contributions, Self Assessment triggers, and tax efficiency tips.

  6. REPORT
     Everything is written to a Markdown (.md) file you can open in VS Code
     (Ctrl+Shift+V for rendered preview) or any Markdown viewer.

REQUIREMENTS
------------
  - Python 3.10+  (https://python.org)
  - LM Studio     (https://lmstudio.ai)  -- free desktop app
      * Load a model (e.g. Meta-Llama-3.1-8B-Instruct-Q4_K_M.gguf)
      * Go to Developer tab (</>)  -> Start Server
      * Server must show: http://127.0.0.1:1234
  - Python libraries (install once):
      pip install pdfplumber requests

USAGE
-----
  Basic (uses ./statements folder by default):
      python tax_agent.py

  Point at your statements folder:
      python tax_agent.py "G:\\My Drive\\Statements"

  Full options:
      python tax_agent.py "G:\\My Drive\\Statements" \
          --model "meta-llama-3.1-8b-instruct" \
          --output "tax_report_2024_25.md"

ARGUMENTS
---------
  pdf_folder   Path to folder containing your PDF bank statements.
               Subfolders are searched automatically. Archive folders and
               statements older than 2024 are skipped.
               Default: ./statements

  --model      The model identifier shown in LM Studio's Developer tab under
               "API Model Identifier". Must match exactly (case-sensitive).
               Default: meta-llama-3.1-8b-instruct

  --output     Filename for the generated Markdown report.
               Default: tax_report.md

WHAT IT LOOKS FOR IN YOUR STATEMENTS
-------------------------------------
  - Salary / employment income (regular large credits from employers)
  - Savings interest (credits labelled "interest", "int", etc.)
  - Dividends (credits labelled "dividend", "div", "DRIP", etc.)
  - Capital gains (large one-off credits from brokers/platforms)
  - ISA deposits (transfers to ISA accounts)
  - Pension contributions (payments to pension providers)
  - Large unexplained credits (anything that does not fit a known category)

LIMITATIONS
-----------
  - The AI can only see what appears in the bank statements. It cannot see:
      * Employer pension contributions (these do not appear in your account)
      * P11D benefits in kind
      * Income taxed at source that does not appear as a bank credit
  - Accuracy depends on how clearly transactions are labelled in your PDFs.
  - This is an ESTIMATE. Always verify with a qualified UK tax adviser.

UPDATING TAX THRESHOLDS
------------------------
  The UK_TAX dictionary near the top of this file contains the 2024/25
  thresholds. When a new tax year starts, update these values. HMRC publishes
  rates at: https://www.gov.uk/income-tax-rates

PRIVACY
-------
  - LM Studio runs entirely offline. No internet connection is used for AI.
  - Your PDFs never leave your computer.
  - The requests library is used only to talk to localhost:1234.

================================================================================
"""

import os
import sys
import json
import re
import argparse
from pathlib import Path
from datetime import datetime


EXTRACTION_LIST_KEYS = [
    "income", "interest", "dividends", "capital_gains",
    "pension_contributions", "isa_deposits", "large_unexplained_credits"
]

EXTRACTION_SCHEMA = {
    "income": {
        "required": ["date", "description", "amount", "type"],
        "numeric": ["amount"],
    },
    "interest": {
        "required": ["date", "bank", "amount", "account_type"],
        "numeric": ["amount"],
    },
    "dividends": {
        "required": ["date", "company", "amount", "is_isa"],
        "numeric": ["amount"],
    },
    "capital_gains": {
        "required": ["date", "asset", "proceeds", "cost", "gain"],
        "numeric": ["proceeds", "cost", "gain"],
    },
    "pension_contributions": {
        "required": ["date", "amount", "type"],
        "numeric": ["amount"],
    },
    "isa_deposits": {
        "required": ["date", "amount", "isa_type"],
        "numeric": ["amount"],
    },
    "large_unexplained_credits": {
        "required": ["date", "amount", "description"],
        "numeric": ["amount"],
    },
}

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# ── Dependencies ──────────────────────────────────────────────────────────────
# pdfplumber : reads PDF files and extracts text and tables
# requests   : makes HTTP calls to the LM Studio local server (localhost only)

try:
    import pdfplumber
except ImportError:
    print("pdfplumber not found -- installing now...")
    os.system(f"{sys.executable} -m pip install pdfplumber -q")
    import pdfplumber

try:
    import requests
except ImportError:
    print("requests not found -- installing now...")
    os.system(f"{sys.executable} -m pip install requests -q")
    import requests


# ── Configuration ─────────────────────────────────────────────────────────────
# These are the two settings most likely to need changing.

# URL of LM Studio's local API server. This does not change unless you have
# manually edited LM Studio's port in its settings.
LM_STUDIO_URL = "http://localhost:1234/v1/chat/completions"

# The model identifier -- must exactly match the string shown in LM Studio's
# Developer tab under "API Model Identifier". If you switch models, update this.
DEFAULT_MODEL = "meta-llama-3.1-8b-instruct"


# ── UK Tax Thresholds 2024/25 ─────────────────────────────────────────────────
# These values are passed directly to the AI so it can calculate liabilities.
# Source: https://www.gov.uk/income-tax-rates
# UPDATE THESE each April when HMRC publishes new rates for the coming year.

UK_TAX = {
    # Income tax bands
    "personal_allowance":       12_570,   # Tax-free income allowance
    "basic_rate_limit":         50_270,   # Income up to here taxed at 20%
    "higher_rate_limit":       125_140,   # Income up to here taxed at 40%
    # Income above 125,140 is taxed at 45% (additional rate)

    # Savings interest allowances (tax-free interest you can earn each year)
    "savings_allowance_basic":   1_000,   # For basic rate (20%) taxpayers
    "savings_allowance_higher":    500,   # For higher rate (40%) taxpayers
    # Additional rate (45%) taxpayers receive no savings allowance

    # Dividends
    "dividend_allowance":          500,   # Tax-free dividend amount per year
    "dividend_basic_rate":       0.0875,  # 8.75%  on dividends above allowance (basic rate)
    "dividend_higher_rate":      0.3375,  # 33.75% on dividends above allowance (higher rate)

    # Capital Gains Tax
    "cgt_allowance":             3_000,   # Tax-free gains per year (Annual Exempt Amount)
    "cgt_basic_rate_shares":      0.10,   # 10% CGT for basic rate taxpayers (shares/funds)
    "cgt_higher_rate_shares":     0.20,   # 20% CGT for higher rate taxpayers (shares/funds)

    # ISA (Individual Savings Account)
    "isa_allowance":            20_000,   # Maximum you can deposit per tax year
    # All income and gains inside an ISA are completely tax-free forever

    # Pension
    "pension_annual_allowance": 60_000,   # Maximum tax-relieved pension contributions per year
}


def parse_tax_year(tax_year: str) -> tuple[int, int]:
    """Parse tax year labels like '2024-25' or '2024/25' into (2024, 2025)."""
    match = re.fullmatch(r"\s*(\d{4})\s*[-/]\s*(\d{2}|\d{4})\s*", tax_year)
    if not match:
        raise ValueError("Use format YYYY-YY or YYYY-YYYY (e.g. 2024-25)")

    start_year = int(match.group(1))
    end_token = match.group(2)
    if len(end_token) == 2:
        end_year = (start_year // 100) * 100 + int(end_token)
    else:
        end_year = int(end_token)

    if end_year != start_year + 1:
        raise ValueError("Tax year must span one year (e.g. 2024-25)")
    return start_year, end_year


def tax_year_label(start_year: int, end_year: int) -> str:
    """Format tax year tuple as YYYY/YY for display."""
    return f"{start_year}/{str(end_year)[-2:]}"


def tax_year_filename_tokens(tax_year: str) -> list[str]:
    """Return filename year tokens for a tax year (e.g. ['2024', '2025'])."""
    start_year, end_year = parse_tax_year(tax_year)
    return [str(start_year), str(end_year)]


def tax_year_date_window(tax_year: str):
    """Return UK tax-year date window as (start_date, end_date)."""
    start_year, end_year = parse_tax_year(tax_year)
    return datetime(start_year, 4, 6).date(), datetime(end_year, 4, 5).date()


def is_iso_date(value: str) -> bool:
    """Validate strict ISO date format YYYY-MM-DD."""
    if not isinstance(value, str) or not DATE_RE.match(value):
        return False
    try:
        datetime.strptime(value, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def parse_iso_date(value: str):
    """Parse ISO date YYYY-MM-DD to date object; return None if invalid."""
    if not is_iso_date(value):
        return None
    return datetime.strptime(value, "%Y-%m-%d").date()


def filter_financial_data_by_tax_year(financial_data: dict, tax_year: str) -> tuple[dict, dict]:
    """Keep only events whose transaction dates fall inside the selected UK tax year."""
    start_date, end_date = tax_year_date_window(tax_year)
    filtered = {k: [] for k in EXTRACTION_LIST_KEYS}
    filtered["notes"] = financial_data.get("notes", "") if isinstance(financial_data.get("notes"), str) else ""

    total_events = 0
    kept_events = 0
    excluded_outside_year = 0
    excluded_invalid_date = 0

    for key in EXTRACTION_LIST_KEYS:
        records = financial_data.get(key, [])
        if not isinstance(records, list):
            continue
        for record in records:
            if not isinstance(record, dict):
                continue
            total_events += 1
            event_date = parse_iso_date(record.get("date", ""))
            if not event_date:
                excluded_invalid_date += 1
                continue
            if start_date <= event_date <= end_date:
                filtered[key].append(record)
                kept_events += 1
            else:
                excluded_outside_year += 1

    filter_note = (
        f"Date filter applied for UK tax year {tax_year}: "
        f"kept {kept_events}/{total_events} events "
        f"({excluded_outside_year} outside range, {excluded_invalid_date} invalid date)."
    )
    filtered["notes"] = (filtered["notes"] + " | " + filter_note).strip(" |")

    stats = {
        "tax_year": tax_year,
        "window_start": str(start_date),
        "window_end": str(end_date),
        "total_events": total_events,
        "kept_events": kept_events,
        "excluded_outside_year": excluded_outside_year,
        "excluded_invalid_date": excluded_invalid_date,
    }
    return filtered, stats


def validate_and_normalize_extraction(raw_text: str) -> tuple[dict | None, str]:
    """Parse, validate, and normalize extraction JSON from the model."""
    cleaned = re.sub(r"```(?:json)?|```", "", raw_text).strip()
    try:
        parsed = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        return None, f"JSON parse error: {exc}"

    if not isinstance(parsed, dict):
        return None, "Top-level response is not a JSON object"

    normalized = {k: [] for k in EXTRACTION_LIST_KEYS}
    normalized["notes"] = ""
    rejected = 0

    for key in EXTRACTION_LIST_KEYS:
        records = parsed.get(key, [])
        if not isinstance(records, list):
            continue
        schema = EXTRACTION_SCHEMA[key]
        for record in records:
            if not isinstance(record, dict):
                rejected += 1
                continue
            if not all(field in record for field in schema["required"]):
                rejected += 1
                continue
            if not is_iso_date(record.get("date")):
                rejected += 1
                continue

            valid = True
            normalized_record = {}
            for field, value in record.items():
                normalized_record[field] = value

            for numeric_field in schema["numeric"]:
                try:
                    normalized_record[numeric_field] = float(record[numeric_field])
                except (TypeError, ValueError):
                    valid = False
                    break

            if "is_isa" in record:
                normalized_record["is_isa"] = bool(record["is_isa"])

            if valid:
                normalized[key].append(normalized_record)
            else:
                rejected += 1

    notes = parsed.get("notes", "")
    normalized["notes"] = notes if isinstance(notes, str) else ""

    total_valid = sum(len(normalized[k]) for k in EXTRACTION_LIST_KEYS)
    if total_valid == 0 and rejected > 0:
        return None, "All extracted records failed schema validation"

    return normalized, ""


def build_batches(statements: dict[str, str], batch_chars: int = 3000, chunk_chars: int = 1800) -> list[str]:
    """Build context-safe batches from full statement text without truncating files."""
    batches = []
    current_batch = ""

    for filename, text in statements.items():
        if not text.strip():
            continue

        parts = [text[i:i + chunk_chars] for i in range(0, len(text), chunk_chars)]
        for part_index, part in enumerate(parts, 1):
            chunk = f"\n=== {filename} (part {part_index}/{len(parts)}) ===\n{part}\n"
            if len(current_batch) + len(chunk) > batch_chars and current_batch:
                batches.append(current_batch)
                current_batch = chunk
            else:
                current_batch += chunk

    if current_batch:
        batches.append(current_batch)
    return batches


def event_signature(event_type: str, event: dict) -> tuple:
    """Build a dedupe signature for extracted events."""
    date = str(event.get("date", "")).strip()
    amount = event.get("amount")
    amount_value = round(float(amount), 2) if isinstance(amount, (int, float)) else None

    if event_type == "income":
        return (date, amount_value, str(event.get("description", "")).strip().lower(), str(event.get("type", "")).strip().lower())
    if event_type == "interest":
        return (date, amount_value, str(event.get("bank", "")).strip().lower(), str(event.get("account_type", "")).strip().lower())
    if event_type == "dividends":
        return (date, amount_value, str(event.get("company", "")).strip().lower(), bool(event.get("is_isa", False)))
    if event_type == "capital_gains":
        return (
            date,
            round(float(event.get("proceeds", 0.0)), 2),
            round(float(event.get("cost", 0.0)), 2),
            round(float(event.get("gain", 0.0)), 2),
            str(event.get("asset", "")).strip().lower(),
        )
    if event_type == "pension_contributions":
        return (date, amount_value, str(event.get("type", "")).strip().lower())
    if event_type == "isa_deposits":
        return (date, amount_value, str(event.get("isa_type", "")).strip().lower())
    return (date, amount_value, str(event.get("description", "")).strip().lower())


# ── PDF Text Extraction ───────────────────────────────────────────────────────

def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Open a single PDF and extract all readable text from it.

    Strategy:
      For each page, first try table extraction (ideal for bank statements
      which are laid out in columns: date | description | debit | credit | balance).
      If no tables are detected, fall back to plain text extraction.

      Table rows are joined with ' | ' so the AI can see the column structure:
        "01 Apr 2024 | BACS CREDIT EMPLOYER LTD | | 2500.00 | 7500.00"

    Returns a single string of all extracted text, or "" if the file cannot
    be read (e.g. it is password protected or a scanned image without OCR).
    """
    text_parts = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"  Reading {Path(pdf_path).name}: {len(pdf.pages)} page(s)")
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    # Table mode: flatten each row into a pipe-separated string
                    for table in tables:
                        for row in table:
                            if row:
                                clean_row = [str(cell).strip() if cell else "" for cell in row]
                                text_parts.append(" | ".join(clean_row))
                else:
                    # Plain text fallback
                    text = page.extract_text()
                    if text:
                        text_parts.append(text)
    except Exception as e:
        print(f"  Warning: Could not read {pdf_path}: {e}")
        return ""
    return "\n".join(text_parts)


def load_all_statements(pdf_folder: str, tax_year: str = "2024-25") -> dict[str, str]:
    """
    Find, filter, deduplicate, and load all relevant PDF statements.

    Steps:
      1. Walk the entire folder tree with rglob() to find every .pdf file.
      2. Resolve each path to its real absolute path and deduplicate.
         Google Drive on Windows can expose the same file at multiple paths
         (via shortcuts or sync folder aliases), causing double-processing.
      3. Skip any files inside folders named 'archive' or 'Archive'.
         These are old historical statements we do not need.
      4. Keep only files whose names contain 2024, 2025, or 2026.
         This focuses the analysis on the current/recent tax years and
         prevents the AI from being overwhelmed by irrelevant old data.
      5. Extract text from each surviving PDF.

    Returns:
      dict mapping relative file path (str) -> extracted text (str)
      e.g. {"Lloyds Statements\\2024_April_Statement.pdf": "01 Apr..."}
    """
    folder = Path(pdf_folder)

    # Step 1 + 2: Walk all subfolders, collect unique PDFs only
    seen_real_paths = set()
    all_pdfs = []
    for p in folder.rglob("*"):
        if p.suffix.lower() == ".pdf":
            real_path = p.resolve()               # Follow symlinks / Drive aliases
            if real_path not in seen_real_paths:  # Only keep first occurrence
                seen_real_paths.add(real_path)
                all_pdfs.append(p)

    if not all_pdfs:
        print(f"\n  No PDF files found in: {pdf_folder}")
        print("    Check the folder path and try again.")
        sys.exit(1)

    # Step 3 + 4: Filter to likely-relevant files for selected tax year
    RELEVANT_YEARS = tax_year_filename_tokens(tax_year)
    SKIP_FOLDERS   = ["archive", "Archive"]    # Folder names to ignore entirely

    relevant_pdfs = []
    skipped_count = 0
    for p in sorted(all_pdfs):
        # Skip if any part of the path is an archive folder
        if any(skip in p.parts for skip in SKIP_FOLDERS):
            skipped_count += 1
            continue
        # Keep if the filename contains a relevant year
        if any(year in p.name for year in RELEVANT_YEARS):
            relevant_pdfs.append(p)
        else:
            skipped_count += 1

    print(f"\n  Found {len(all_pdfs)} total PDFs.")
    print(f"  Using  {len(relevant_pdfs)} likely relevant to tax year {tax_year}.")
    print(f"  Skipped {skipped_count} (archives / old statements).\n")

    # Step 5: Extract text from each relevant PDF
    statements = {}
    for pdf_path in relevant_pdfs:
        try:
            rel_path = pdf_path.relative_to(folder)  # e.g. Lloyds\2024_April.pdf
        except ValueError:
            rel_path = pdf_path.name
        text = extract_text_from_pdf(str(pdf_path))
        if text.strip():                              # Only store files with readable text
            statements[str(rel_path)] = text

    return statements


# ── LM Studio Connection ──────────────────────────────────────────────────────

def check_lm_studio() -> bool:
    """
    Ping LM Studio's /v1/models endpoint to confirm the server is up.
    Returns True if reachable and responding, False otherwise.
    Times out after 3 seconds to fail fast if LM Studio is not running.
    """
    try:
        r = requests.get("http://localhost:1234/v1/models", timeout=3)
        return r.status_code == 200
    except Exception:
        return False


def ask_llm(system_prompt: str, user_message: str, model: str = DEFAULT_MODEL) -> str:
    """
    Send a prompt to the LM Studio local server and return the AI's response.

    LM Studio exposes an OpenAI-compatible API at /v1/chat/completions.
    We send a 'messages' array with two roles:
      - system : sets the AI's persona and strict behaviour rules
      - user   : the actual data or question to process

    Key settings:
      temperature=0.1  Makes the AI deterministic and precise rather than
                       creative. Essential for financial data -- we want
                       "£1,247.50" not "approximately a thousand pounds".
      max_tokens=4096  Maximum length of the AI's reply.
      stream=False     Receive the complete response at once rather than
                       as a stream of tokens.

    The API returns JSON; we extract choices[0].message.content which is
    the AI's actual reply text.

    On connection failure, we print a helpful error and exit immediately
    rather than silently producing an empty report.
    """
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": user_message}
        ],
        "temperature": 0.1,    # Low = consistent and precise
        "max_tokens":  4096,   # Max response length in tokens
        "stream":      False   # Full response, not streamed
    }
    try:
        response = requests.post(LM_STUDIO_URL, json=payload, timeout=300)
        response.raise_for_status()  # Raises HTTPError on 4xx/5xx status codes
        return response.json()["choices"][0]["message"]["content"]
    except requests.exceptions.ConnectionError:
        print("\n  Cannot connect to LM Studio. Is the server running?")
        print("  Fix: Developer tab (</>) -> toggle 'Start Server' ON")
        sys.exit(1)
    except KeyError:
        # The response came back but in an unexpected shape
        return f"Unexpected API response format: {response.json()}"
    except Exception as e:
        return f"Error calling LM Studio: {e}"


# ── System Prompts ────────────────────────────────────────────────────────────
# A "system prompt" is the instruction we give the AI before each conversation
# to set its role and rules. Keeping these short and specific gets better
# results from smaller local models like Llama 3.1 8B.

EXTRACTION_SYSTEM = """You are a precise financial data extractor.
Extract structured financial data from bank statements.
Respond ONLY with valid JSON -- no explanation, no markdown fences, no preamble."""

ANALYSIS_SYSTEM = """You are a UK tax adviser with expertise in personal tax.
Provide accurate, conservative estimates based on the data provided.
Always flag uncertainty and recommend professional advice for final filing.
Use UK tax rates for the 2024/25 tax year."""


# ── Batch Extraction ──────────────────────────────────────────────────────────
# WHY BATCHING?
# The Llama 3.1 8B model has a context window of ~4096 tokens (~3000 words).
# Sending all statements at once (potentially 100,000+ characters) causes a
# 400 Bad Request error. The solution is to split into small batches, extract
# from each, then merge all results together at the end.

# The JSON schema we ask the AI to populate for each batch.
# Double braces {{ }} are Python's way of writing literal { } inside an f-string.
EXTRACTION_PROMPT = """From this bank statement data extract ALL financial events relevant to UK tax.
Return ONLY a JSON object. No explanation, no markdown, no preamble. Just the JSON.

Required structure:
{{"income":[{{"date":"YYYY-MM-DD","description":"...","amount":0.00,"type":"salary|freelance|rental|other"}}],
"interest":[{{"date":"YYYY-MM-DD","bank":"...","amount":0.00,"account_type":"savings|current|ISA|other"}}],
"dividends":[{{"date":"YYYY-MM-DD","company":"...","amount":0.00,"is_isa":false}}],
"capital_gains":[{{"date":"YYYY-MM-DD","asset":"...","proceeds":0.00,"cost":0.00,"gain":0.00}}],
"pension_contributions":[{{"date":"YYYY-MM-DD","amount":0.00,"type":"personal|employer|SIPP"}}],
"isa_deposits":[{{"date":"YYYY-MM-DD","amount":0.00,"isa_type":"cash|stocks|LISA|IFISA"}}],
"large_unexplained_credits":[{{"date":"YYYY-MM-DD","amount":0.00,"description":"..."}}],
"notes":"any observations about data quality or missing information"}}

Statement data:
"""

# Fallback empty result returned if every batch fails to parse
EMPTY_EXTRACTION = {
    "income": [], "interest": [], "dividends": [], "capital_gains": [],
    "pension_contributions": [], "isa_deposits": [], "large_unexplained_credits": [],
    "notes": "No data could be extracted."
}


def merge_extractions(batch_results: list[dict]) -> dict:
    """
    Combine multiple batch JSON results into one master dictionary.

    Each batch produces its own JSON object with lists of financial events.
    This function concatenates all the lists so we get a single complete
    picture of all events across all statements and batches.

    Example:
      Batch 1 found: 3 interest payments, 1 salary credit
      Batch 2 found: 5 interest payments, 2 ISA deposits
      Merged result: 8 interest payments, 1 salary credit, 2 ISA deposits
    """
    merged = {k: [] for k in EXTRACTION_LIST_KEYS}
    merged["notes"] = ""
    notes_list = []
    seen_by_key = {k: set() for k in EXTRACTION_LIST_KEYS}

    for result in batch_results:
        for key in EXTRACTION_LIST_KEYS:
            if isinstance(result.get(key), list):
                for event in result[key]:
                    if not isinstance(event, dict):
                        continue
                    signature = event_signature(key, event)
                    if signature in seen_by_key[key]:
                        continue
                    seen_by_key[key].add(signature)
                    merged[key].append(event)
        if result.get("notes"):
            notes_list.append(result["notes"])

    merged["notes"] = " | ".join(notes_list)
    return merged


def extract_financial_data(statements: dict[str, str], model: str) -> dict:
    """
    Core extraction function -- splits statements into batches and processes each.

    Batch construction (greedy bin-packing):
      We iterate through statements, building a batch string. When adding the
      next file would push the batch over BATCH_CHARS, we save the current
      batch and start a new one. Each file's text is truncated to 2000 chars
      -- we only need enough to see the key transactions, not every word.

    For each batch:
      1. Prepend the extraction prompt (the JSON schema instructions)
      2. Send to LM Studio via ask_llm()
      3. Strip any markdown fences the AI may have added despite instructions
      4. Parse the JSON response
      5. Store the result, or log a parse error and continue

    Finally, merge_extractions() combines all batch results into one dict.
    """
    print("\nStep 1: Extracting financial data from statements...")

    BATCH_CHARS = 3000
    CHUNK_CHARS = 1800
    batches = build_batches(statements, batch_chars=BATCH_CHARS, chunk_chars=CHUNK_CHARS)

    print(f"  Sending {len(statements)} files as {len(batches)} batches to the AI...")

    # Process each batch sequentially
    all_results = []
    for i, batch in enumerate(batches, 1):
        print(f"  Batch {i}/{len(batches)}...", end=" ", flush=True)

        base_prompt = EXTRACTION_PROMPT + batch
        validated = None
        last_error = ""
        max_attempts = 3

        for attempt in range(1, max_attempts + 1):
            prompt = base_prompt
            if attempt > 1:
                prompt += (
                    "\n\nIMPORTANT: Previous response was invalid. "
                    "Return ONLY valid JSON exactly matching the required structure and ISO dates (YYYY-MM-DD)."
                )
            raw = ask_llm(EXTRACTION_SYSTEM, prompt, model)
            validated, last_error = validate_and_normalize_extraction(raw)
            if validated is not None:
                break

        if validated is not None:
            all_results.append(validated)
            event_count = sum(len(validated[k]) for k in EXTRACTION_LIST_KEYS)
            print(f"found {event_count} financial events")
        else:
            print(f"validation failed after {max_attempts} attempts ({last_error}), skipping")

    if not all_results:
        print("  Warning: No data extracted from any batch.")
        print("  Try checking LM Studio logs for errors.")
        return EMPTY_EXTRACTION

    # Combine all batch results into one master dictionary
    merged = merge_extractions(all_results)
    total = sum(len(v) for v in merged.values() if isinstance(v, list))
    print(f"\n  Done. {total} financial events extracted across all statements.")
    return merged


# ── Tax Analysis ──────────────────────────────────────────────────────────────

def calculate_tax_implications(financial_data: dict, model: str, tax_year: str = "2024-25") -> str:
    """
    Send the merged financial data and UK tax thresholds to the AI for analysis.

    By this point, the financial data is much smaller than the raw PDFs were --
    it is a structured JSON summary of events rather than pages of raw text.
    This fits comfortably in one prompt.

    We explicitly tell the AI which sections to cover so the report is
    consistent and complete. The AI's response comes back as Markdown-formatted
    text which is embedded directly into the final report file.
    """
    print("\nStep 2: Calculating UK tax implications...")

    tax_context = json.dumps(UK_TAX, indent=2)
    data_str    = json.dumps(financial_data, indent=2)

    prompt = f"""Using the extracted financial data and UK {tax_year} tax thresholds below,
provide a comprehensive UK tax analysis.

UK TAX THRESHOLDS {tax_year}:
{tax_context}

EXTRACTED FINANCIAL DATA:
{data_str}

Please analyse and report on each of the following sections:

1. INCOME TAX
   - Total income identified
   - Estimated tax band (basic / higher / additional rate)
   - Estimated income tax liability
   - Salary/PAYE already deducted (if visible in statements)

2. SAVINGS INTEREST
   - Total interest earned across all accounts
   - Amount covered by the Personal Savings Allowance
   - Taxable interest remaining (if any)
   - Note: ISA interest is completely tax-free

3. DIVIDENDS
   - Total dividends received outside an ISA
   - Amount covered by the £500 dividend allowance
   - Taxable dividends and estimated tax at the relevant rate

4. CAPITAL GAINS TAX
   - Total gains identified
   - Amount covered by the £3,000 Annual Exempt Amount
   - Taxable gain and estimated CGT at the relevant rate

5. ISA USAGE
   - Total deposited into ISAs this tax year
   - Remaining allowance of the £20,000 annual limit
   - Estimated income and gains sheltered tax-free

6. PENSION CONTRIBUTIONS
   - Personal contributions identified
   - Estimated tax relief received
   - Whether within the £60,000 annual allowance

7. SELF-ASSESSMENT TRIGGERS
   - Whether a Self Assessment return is likely required
   - Key reasons (e.g. income over £100k, untaxed income over £1k, CGT events)

8. TAX EFFICIENCY OPPORTUNITIES
   - Specific actionable suggestions based on this person's actual data
   - For example: top up ISA, increase pension contributions before year end

9. SUMMARY TABLE
   - Estimated total tax liability broken down by category
   - Key HMRC deadlines to be aware of

10. CAVEATS
    - What this analysis cannot see (employer pension, P11D, etc.)
    - Strong recommendation to consult a professional tax adviser
    - Reminder that this is an estimate only

Be specific with amounts where the data allows. Flag all assumptions clearly."""

    return ask_llm(ANALYSIS_SYSTEM, prompt, model)


# ── Report Generation ─────────────────────────────────────────────────────────

def generate_report(financial_data: dict, tax_analysis: str, output_path: str, tax_year: str = "2024-25", filter_stats: dict | None = None):
    """
    Write the complete tax report to a Markdown (.md) file.

    The report has three main sections:
      1. Header, timestamp, and disclaimer
      2. Raw extracted financial data as JSON -- lets you verify the AI read
         your statements correctly before trusting the analysis
      3. The AI's full tax analysis in plain English Markdown

    Saved as UTF-8 to handle pound signs (pound) and other non-ASCII characters
    without errors on Windows (which defaults to cp1252 encoding).

    To view the rendered Markdown in VS Code:
      Open the file -> press Ctrl+Shift+V for a live rendered preview
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    date_filter_section = ""
    if filter_stats:
        date_filter_section = f"""
## Tax-Year Date Filter
Applied window: {filter_stats['window_start']} to {filter_stats['window_end']} (UK tax year {tax_year})

- Total extracted events: {filter_stats['total_events']}
- Included in tax year: {filter_stats['kept_events']}
- Excluded (outside year): {filter_stats['excluded_outside_year']}
- Excluded (invalid date): {filter_stats['excluded_invalid_date']}

---

"""

    report = f"""# UK Tax Analysis Report
Generated: {now}
CONFIDENTIAL - Processed entirely on your local machine. No data was sent externally.

---

## Disclaimer
This report is generated by an AI assistant using data extracted from your bank statements.
It is an ESTIMATE ONLY and should not be used as a substitute for professional tax advice.
Always consult a qualified UK tax adviser or accountant before filing with HMRC.

---

## Extracted Financial Data
The AI identified the following financial events from your PDF statements.
**Please review this section carefully.** If key transactions are missing or
incorrectly classified, the tax analysis below will also be affected.

```json
{json.dumps(financial_data, indent=2)}
```

---

{date_filter_section}

---

## Tax Analysis and Implications

{tax_analysis}

---

## Next Steps

1. **Verify the extracted data** above against your actual statements
2. **Gather supporting documents**: P60, P11D, dividend vouchers, broker contract notes
3. **Register for Self Assessment** if required: https://www.gov.uk/register-for-self-assessment
4. **Filing deadlines**: 31 October (paper return) | 31 January (online return)
5. **Consult a qualified tax professional** before submitting anything to HMRC

---
*Report generated locally using LM Studio + Meta Llama 3.1 8B Instruct.*
*Your financial data never left your computer.*
"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"\n  Report saved to: {output_path}")


# ── Entry Point ───────────────────────────────────────────────────────────────

def main():
    """
    Command-line entry point.

    Parses arguments then orchestrates the full pipeline:
      check_lm_studio()
        -> load_all_statements()
          -> extract_financial_data()
            -> calculate_tax_implications()
              -> generate_report()

    The script exits with a clear error message if LM Studio is not running
    or if no PDFs are found, rather than silently producing an empty report.
    """
    parser = argparse.ArgumentParser(
        description="UK Tax Analysis Agent -- reads bank statement PDFs and analyses "
                    "tax implications locally using LM Studio. No data leaves your machine."
    )
    parser.add_argument(
        "pdf_folder",
        nargs="?",
        default="./statements",
        help="Path to folder containing your PDF bank statements (default: ./statements)"
    )
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=f"LM Studio model identifier shown in the Developer tab (default: {DEFAULT_MODEL})"
    )
    parser.add_argument(
        "--output",
        default="tax_report.md",
        help="Output report filename (default: tax_report.md)"
    )
    parser.add_argument(
        "--tax-year",
        default="2024-25",
        help="Tax year label for filtering/reporting in format YYYY-YY (default: 2024-25)"
    )
    args = parser.parse_args()

    try:
        start_year, end_year = parse_tax_year(args.tax_year)
        normalized_tax_year = f"{start_year}-{str(end_year)[-2:]}"
    except ValueError as exc:
        print(f"\n  Invalid --tax-year value '{args.tax_year}': {exc}")
        sys.exit(1)

    print("=" * 60)
    print("  UK Tax Analysis Agent - Local and Private")
    print("=" * 60)
    print(f"  Backend : LM Studio (localhost:1234)")
    print(f"  Model   : {args.model}")
    print(f"  Tax year: {tax_year_label(start_year, end_year)}")
    print(f"  Folder  : {args.pdf_folder}")
    print(f"  Output  : {args.output}")

    # Confirm LM Studio is reachable before doing any work
    print("\n  Checking LM Studio connection...")
    if not check_lm_studio():
        print("\n  LM Studio server is not running!")
        print("  Fix: Open LM Studio -> Developer tab -> Start Server")
        sys.exit(1)
    print("  LM Studio is running.")

    # Load and filter PDFs from the statements folder
    statements = load_all_statements(args.pdf_folder, tax_year=normalized_tax_year)
    print(f"  Loaded {len(statements)} statement(s) with readable text.")

    # Extract structured financial events from all statements (batched)
    financial_data = extract_financial_data(statements, args.model)

    # Apply strict transaction-date filtering for the selected UK tax year
    financial_data, date_filter_stats = filter_financial_data_by_tax_year(financial_data, normalized_tax_year)
    print(
        "  Date filter "
        f"({date_filter_stats['window_start']} to {date_filter_stats['window_end']}): "
        f"kept {date_filter_stats['kept_events']}/{date_filter_stats['total_events']} events "
        f"({date_filter_stats['excluded_outside_year']} outside year, "
        f"{date_filter_stats['excluded_invalid_date']} invalid date)."
    )

    # Run UK tax analysis against tax-year-filtered data
    tax_analysis = calculate_tax_implications(financial_data, args.model, tax_year=normalized_tax_year)

    # Write the full report to disk
    generate_report(
        financial_data,
        tax_analysis,
        args.output,
        tax_year=normalized_tax_year,
        filter_stats=date_filter_stats,
    )

    # Also print the analysis to the terminal for a quick read
    print("\n" + "=" * 60)
    print("  TAX ANALYSIS SUMMARY")
    print("=" * 60)
    print(tax_analysis)
    print("\n" + "=" * 60)
    print("  Done. Always verify with a qualified UK tax professional.")
    print("=" * 60)


# Only run main() if this script is executed directly.
# If someone imports this file as a module, main() will not auto-run.
if __name__ == "__main__":
    main()
