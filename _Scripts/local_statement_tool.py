"""
Local-only bank statement parsing and spending analysis tool.

Security model:
- Runs entirely on the local machine.
- Reads only local files supplied by the user.
- Does not call cloud APIs, web services, AI/LLM APIs, telemetry, or upload endpoints.
- Does not transmit statement content outside this process.

Dependencies:
- pypdf
- openpyxl
- Python standard library modules only
"""

import argparse
import csv
import html
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from pypdf import PdfReader
from pypdf.errors import PdfReadError


# Deterministic category list used in reports and policy validation.
ESSENTIAL_CATEGORIES = [
    "Mortgage/Rent",
    "Loans/Credit Cards/Car Lease/Other liabilities",
    "Council Tax",
    "Gas/Electricity",
    "Broadband/Phone landline",
    "Water",
    "Mobile Phone",
    "TV/Satellite - Sky etc",
    "Memberships/Club fees",
    "House maintenance",
    "Food/Shopping",
    "Car /Travel Expenses/Petrol",
    "Life Cover/Income Protection etc",
    "Home Insurance",
    "Car Insurance",
    "Pet Insurance",
    "Regular Savings - ISAs, Endowments etc",
    "Other Living Expenses",
]

DISCRETIONARY_CATEGORIES = [
    "Gym",
    "Holidays/Trips",
    "Entertainment - restaurant/cinema/pub etc",
    "Clothes",
    "Health & Beauty - Hair etc",
    "Pastimes/Sport/Hobbies expenses",
    "Monetary Gifts to family",
    "Other",
]

ALL_CATEGORIES = ESSENTIAL_CATEGORIES + DISCRETIONARY_CATEGORIES
DEFAULT_POLICY_JSON = Path(__file__).with_name("expense_policy_hints.json")
DEFAULT_OUTPUT_FOLDER = r"G:\My Drive\Statements\Analysis Outputs\Local Statement Tool"

BUILTIN_KEYWORD_RULES = [
    ("Mortgage/Rent", [r"\bRENT\b", r"\bMORTGAGE\b", r"\bGROUND RENT\b"]),
    (
        "Loans/Credit Cards/Car Lease/Other liabilities",
        [r"\bCAPITAL ONE\b", r"\bBARCLAYCARD\b", r"\bMBNA\b", r"\bLOAN\b", r"\bFINANCE\b"],
    ),
    ("Council Tax", [r"\bCOUNCIL\b"]),
    ("Gas/Electricity", [r"\bEDF\b", r"\bBRITISH GAS\b", r"\bE\.?ON\b", r"\bOCTOPUS\b", r"\bFUSE ENERGY\b"]),
    ("Broadband/Phone landline", [r"\bVIRGIN MEDIA\b", r"\bBT\b", r"\bTALKTALK\b", r"\bSKY BROADBAND\b"]),
    ("Water", [r"\bUNITED UTILITIES\b", r"\bWATER\b"]),
    ("Mobile Phone", [r"\bVODAFONE\b", r"\bO2\b", r"\bEE\b", r"\bGIFFGAFF\b", r"\bTHREE\b"]),
    ("TV/Satellite - Sky etc", [r"\bSKY\b", r"\bNETFLIX\b", r"\bDISNEY\b", r"\bPRIME VIDEO\b"]),
    ("Memberships/Club fees", [r"\bUNISON\b", r"\bMEMBERSHIP\b", r"\bRAMBLERS\b", r"\bPRIME\b"]),
    ("House maintenance", [r"\bB&Q\b", r"\bWICKES\b", r"\bHOMEBASE\b", r"\bSCREWFIX\b", r"\bCLEANING\b"]),
    ("Food/Shopping", [r"\bTESCO\b", r"\bASDA\b", r"\bALDI\b", r"\bLIDL\b", r"\bSAINSBURY\b", r"\bMORRISONS\b"]),
    ("Car /Travel Expenses/Petrol", [r"\bSHELL\b", r"\bESSO\b", r"\bBP\b", r"\bTEXACO\b", r"\bNCP\b", r"\bPARKING\b"]),
    ("Home Insurance", [r"\bHOME INSURANCE\b", r"\bCONTENTS INSURANCE\b"]),
    ("Car Insurance", [r"\bCAR INSURANCE\b", r"\bCHURCHILL\b", r"\bHASTINGS\b"]),
    ("Pet Insurance", [r"\bPET INSURANCE\b", r"\bLV PET\b", r"\bPETPLAN\b"]),
    ("Regular Savings - ISAs, Endowments etc", [r"\bISA\b", r"\bTRADING 212\b", r"\bVANGUARD\b", r"\bSAVINGS\b"]),
    ("Gym", [r"\bGYM\b", r"\bJD GYMS\b", r"\bPUREGYM\b"]),
    ("Holidays/Trips", [r"\bJET2\b", r"\bTUI\b", r"\bAIRBNB\b", r"\bBOOKING\b", r"\bHOTEL\b", r"\bRYANAIR\b"]),
    ("Entertainment - restaurant/cinema/pub etc", [r"\bCINEMA\b", r"\bPUB\b", r"\bRESTAURANT\b", r"\bSKIDDLE\b"]),
    ("Clothes", [r"\bADIDAS\b", r"\bNIKE\b", r"\bPRIMARK\b", r"\bASOS\b", r"\bDECKERSEUR\b"]),
    ("Health & Beauty - Hair etc", [r"\bSPECSAVERS\b", r"\bBOOTS\b", r"\bSUPERDRUG\b", r"\bHAIR\b", r"\bBARBER\b"]),
    ("Pastimes/Sport/Hobbies expenses", [r"\bPLAYSTATION\b", r"\bXBOX\b", r"\bSTEAM\b", r"\bGOOGLE PLAY\b"]),
]

SECTION_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SECTION_HEADER_FONT = Font(color="FFFFFF", bold=True)
TABLE_HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

PRIVACY_AUDIT = [
    ("pypdf", "Third-party", "No", "Reads local PDF files; no network calls in this script."),
    ("openpyxl", "Third-party", "No", "Writes local Excel workbooks; no network calls in this script."),
    ("argparse", "Stdlib", "No", "Parses CLI arguments locally."),
    ("csv", "Stdlib", "No", "Reads/writes local CSV files only."),
    ("html", "Stdlib", "No", "Escapes text for local HTML report generation."),
    ("json", "Stdlib", "No", "Reads local JSON policy file only."),
    ("re", "Stdlib", "No", "Regex matching in local process memory."),
    ("collections", "Stdlib", "No", "In-memory aggregations only."),
    ("dataclasses", "Stdlib", "No", "Local data structure definitions."),
    ("datetime", "Stdlib", "No", "Date parsing/formatting only."),
    ("decimal", "Stdlib", "No", "Exact decimal arithmetic only."),
    ("pathlib", "Stdlib", "No", "Local filesystem path handling."),
    ("typing", "Stdlib", "No", "Type hints only."),
]


@dataclass
class Transaction:
    date: datetime
    description: str
    tx_type: str
    money_in: Decimal
    money_out: Decimal
    balance: Optional[Decimal]
    source_bank: str
    source_file: str
    category: str = ""
    category_reason: str = ""


def parse_decimal(value: str) -> Decimal:
    cleaned = (value or "").strip()
    if not cleaned:
        return Decimal("0")
    cleaned = (
        cleaned.replace(",", "")
        .replace("£", "")
        .replace("?", "")
        .replace("\u00A3", "")
    )
    if cleaned.lower() == "blank":
        return Decimal("0")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return Decimal("0")


def parse_date(value: str) -> Optional[datetime]:
    for fmt in (
        "%d %b %Y",
        "%d %B %Y",
        "%d %b %y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(value.strip(), fmt)
        except ValueError:
            continue
    return None


def format_currency(value: Decimal) -> str:
    return f"£{value:.2f}"


def extract_pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def detect_pdf_provider(text: str) -> str:
    upper = text.upper()
    if "LLOYDS BANK" in upper and "YOUR TRANSACTIONS" in upper:
        return "lloyds"
    if "CHASE IS A REGISTERED TRADEMARK" in upper or "ACCOUNT STATEMENT" in upper:
        return "chase"
    if "CAPITAL ONE" in upper and "YOUR TRANSACTION DETAILS" in upper:
        return "capital_one"
    return "unknown"


def parse_lloyds_pdf(text: str, source_file: str) -> list[Transaction]:
    pattern = re.compile(
        r"Date\s*\n(?P<date>\d{2} [A-Za-z]{3} \d{2})\.\s*\n"
        r"Description\s*\n(?P<description>.*?)\.\s*\n"
        r"Type\s*\n(?P<transaction_type>.*?)\.\s*\n"
        r"Money In .*?\n(?P<money_in>.*?)\.\s*\n"
        r"Money Out .*?\n(?P<money_out>.*?)\.\s*\n"
        r"Balance .*?\n(?P<balance>[\d,]+\.\d{2})\.",
        re.DOTALL,
    )
    rows = []
    for match in pattern.finditer(text):
        dt = parse_date(match.group("date"))
        if not dt:
            continue
        rows.append(
            Transaction(
                date=dt,
                description=match.group("description").strip(),
                tx_type=match.group("transaction_type").strip(),
                money_in=parse_decimal(match.group("money_in")),
                money_out=parse_decimal(match.group("money_out")),
                balance=parse_decimal(match.group("balance")),
                source_bank="Lloyds",
                source_file=source_file,
            )
        )
    return rows


def parse_chase_pdf(text: str, source_file: str) -> list[Transaction]:
    pattern = re.compile(
        r"(?P<date>\d{2} [A-Za-z]{3} \d{4})\s+"
        r"(?P<description>.+?)\n"
        r"(?P<transaction_type>Transfer|Payment|Purchase|Interest)\s*\n"
        r"(?P<signed_amount>[+-][£\u00A3?]?[\d,]+\.\d{2})\s+"
        r"(?P<balance>[£\u00A3?]?[\d,]+\.\d{2})",
        re.DOTALL,
    )
    rows = []
    for match in pattern.finditer(text):
        dt = parse_date(match.group("date"))
        if not dt:
            continue
        description = match.group("description").replace("\n", " ").strip()
        if description.lower().endswith("opening balance") or description.lower().endswith("closing balance"):
            continue
        signed = parse_decimal(match.group("signed_amount"))
        money_in = signed if signed > 0 else Decimal("0")
        money_out = -signed if signed < 0 else Decimal("0")
        rows.append(
            Transaction(
                date=dt,
                description=description,
                tx_type=match.group("transaction_type").strip(),
                money_in=money_in,
                money_out=money_out,
                balance=parse_decimal(match.group("balance")),
                source_bank="Chase",
                source_file=source_file,
            )
        )
    return rows


def parse_capital_one_pdf(text: str, source_file: str) -> list[Transaction]:
    statement_date_match = re.search(
        r"Statement date\s*\n(?P<statement_date>\d{1,2} [A-Za-z]+ \d{2})",
        text,
        re.IGNORECASE,
    )
    if not statement_date_match:
        return []

    statement_date = datetime.strptime(statement_date_match.group("statement_date"), "%d %B %y")
    day_month_pattern = re.compile(r"^\d{2} [A-Za-z]{3}$")
    amount_pattern = re.compile(r"^[\d,]+\.\d{2}$")
    month_lookup = {
        "JAN": 1,
        "FEB": 2,
        "MAR": 3,
        "APR": 4,
        "MAY": 5,
        "JUN": 6,
        "JUL": 7,
        "AUG": 8,
        "SEP": 9,
        "OCT": 10,
        "NOV": 11,
        "DEC": 12,
    }
    rows: list[Transaction] = []
    lines = [line.strip() for line in text.splitlines()]

    i = 0
    while i < len(lines):
        if not day_month_pattern.match(lines[i]):
            i += 1
            continue
        day_month = lines[i]
        desc_lines = []
        amount = ""
        j = i + 1
        while j < len(lines):
            candidate = lines[j].strip()
            if not candidate:
                j += 1
                continue
            if day_month_pattern.match(candidate) or candidate == "STATEMENT TOTALS":
                break
            if amount_pattern.match(candidate):
                amount = candidate
                j += 1
                break
            if candidate in {"ENG", "GB", "GBR", "Paid in", "Paid out", "Your transaction details"}:
                j += 1
                continue
            if candidate.lower().startswith("on ") or candidate.startswith("Page "):
                j += 1
                continue
            desc_lines.append(candidate)
            j += 1

        if amount:
            month = month_lookup[day_month.split()[1].upper()]
            year = statement_date.year - 1 if month > statement_date.month else statement_date.year
            dt = datetime.strptime(f"{day_month} {year}", "%d %b %Y")
            description = " ".join(desc_lines).strip()
            if description and "STATEMENT TOTALS" not in description:
                is_credit = "PAYMENT RECEIVED" in description.upper() or "REFUND" in description.upper()
                rows.append(
                    Transaction(
                        date=dt,
                        description=description,
                        tx_type="CRD" if is_credit else "DEB",
                        money_in=parse_decimal(amount) if is_credit else Decimal("0"),
                        money_out=Decimal("0") if is_credit else parse_decimal(amount),
                        balance=None,
                        source_bank="Capital One",
                        source_file=source_file,
                    )
                )
        i = j
    return rows


def parse_csv_file(path: Path) -> list[Transaction]:
    rows: list[Transaction] = []
    with path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames:
            return rows
        columns = {c.lower().strip(): c for c in reader.fieldnames}

        def find_col(options):
            for opt in options:
                for col_lower, original in columns.items():
                    if opt in col_lower:
                        return original
            return None

        date_col = find_col(["date", "transaction date"])
        desc_col = find_col(["description", "details", "merchant"])
        type_col = find_col(["type"])
        in_col = find_col(["money in", "credit", "paid in", "amount in"])
        out_col = find_col(["money out", "debit", "paid out", "amount out"])
        amount_col = find_col(["amount"]) if not (in_col and out_col) else None
        balance_col = find_col(["balance"])

        for record in reader:
            date_text = (record.get(date_col or "", "") or "").strip()
            dt = parse_date(date_text)
            if not dt:
                continue
            amount = parse_decimal(record.get(amount_col or "", "")) if amount_col else Decimal("0")
            money_in = parse_decimal(record.get(in_col or "", "")) if in_col else (amount if amount > 0 else Decimal("0"))
            money_out = parse_decimal(record.get(out_col or "", "")) if out_col else (-amount if amount < 0 else Decimal("0"))
            rows.append(
                Transaction(
                    date=dt,
                    description=(record.get(desc_col or "", "") or "").strip(),
                    tx_type=(record.get(type_col or "", "") or "").strip(),
                    money_in=money_in,
                    money_out=money_out,
                    balance=parse_decimal(record.get(balance_col or "", "")) if balance_col else None,
                    source_bank="CSV Import",
                    source_file=path.name,
                )
            )
    return rows


def load_policy(policy_path: Path):
    normalize_patterns = [r"\bPAYPAL\b", r"\bCARD PAYMENT TO\b", r"\bCARD PURCHASE\b", r"\bDEBIT CARD\b", r"\bVISA\b"]
    policy_hints = []
    if not policy_path.exists():
        return normalize_patterns, policy_hints
    policy_data = json.loads(policy_path.read_text(encoding="utf-8"))
    normalize_patterns = policy_data.get("normalize_remove_patterns", normalize_patterns)
    policy_hints = policy_data.get("policy_hints", [])
    return normalize_patterns, policy_hints


def normalize_for_matching(text: str, normalize_patterns: list[str]) -> str:
    normalized = (text or "").upper()
    for pattern in normalize_patterns:
        normalized = re.sub(pattern, " ", normalized)
    normalized = re.sub(r"[^A-Z0-9]+", " ", normalized)
    return re.sub(r"\s+", " ", normalized).strip()


def categorize_transaction(transaction: Transaction, normalize_patterns, policy_hints):
    normalized_desc = normalize_for_matching(transaction.description, normalize_patterns)

    for hint in policy_hints:
        tokens = hint.get("tokens", [])
        if tokens and all(token.upper() in normalized_desc for token in tokens):
            return hint.get("category", "Other"), hint.get("reason", "Policy hint matched.")

    for category, patterns in BUILTIN_KEYWORD_RULES:
        for pattern in patterns:
            matched = re.search(pattern, normalized_desc)
            if matched:
                return category, f"Keyword rule matched '{matched.group(0)}'."

    if transaction.tx_type.upper() == "DD":
        return "Other Living Expenses", "Fallback DD rule."
    return "Other", "Fallback rule (no keyword match)."


def normalize_merchant_name(description: str, normalize_patterns: list[str]) -> str:
    normalized = normalize_for_matching(description, normalize_patterns)
    return normalized[:40] if normalized else "UNKNOWN"


def is_non_transaction_row(tx: Transaction) -> bool:
    text = (tx.description or "").upper()
    tx_type = tx.tx_type.upper()
    if tx_type in {"TRANSFER", "INTEREST", "TFR", "FPO", "DEP"}:
        if any(token in text for token in ("HORSFIELD", "CRAIG", "NOTTINGHAM", "TO CRAIG", "FROM CRAIG", "CRAIG'S ACCOUNT")):
            return True
        if tx_type in {"TRANSFER", "INTEREST"}:
            return True
    if "OPENING BALANCE" in text or "CLOSING BALANCE" in text or "BALANCE TRANSFER" in text:
        return True
    return False


def analyse_transactions(transactions: list[Transaction], normalize_patterns, policy_hints):
    filtered_transactions = [tx for tx in transactions if not is_non_transaction_row(tx)]

    for tx in filtered_transactions:
        tx.category, tx.category_reason = categorize_transaction(tx, normalize_patterns, policy_hints)

    income_total = sum((tx.money_in for tx in filtered_transactions), Decimal("0"))
    spending_total = sum((tx.money_out for tx in filtered_transactions), Decimal("0"))

    monthly = defaultdict(lambda: {"income": Decimal("0"), "spending": Decimal("0"), "net": Decimal("0")})
    category_spend = defaultdict(Decimal)
    merchant_spend = defaultdict(Decimal)
    recurring_groups = defaultdict(list)

    for tx in filtered_transactions:
        month_key = tx.date.strftime("%Y-%m")
        monthly[month_key]["income"] += tx.money_in
        monthly[month_key]["spending"] += tx.money_out
        monthly[month_key]["net"] = monthly[month_key]["income"] - monthly[month_key]["spending"]

        if tx.money_out > 0:
            category_spend[tx.category] += tx.money_out
            merchant = normalize_merchant_name(tx.description, normalize_patterns)
            merchant_spend[merchant] += tx.money_out
            recurring_groups[(merchant, tx.money_out)].append(tx.date)

    monthly_rows = []
    running_net = Decimal("0")
    for month_key in sorted(monthly.keys()):
        running_net += monthly[month_key]["net"]
        monthly_rows.append(
            {
                "month": month_key,
                "income": monthly[month_key]["income"],
                "spending": monthly[month_key]["spending"],
                "net": monthly[month_key]["net"],
                "running_net": running_net,
            }
        )

    category_rows = sorted(
        [{"category": cat, "spending": amt} for cat, amt in category_spend.items()],
        key=lambda row: row["spending"],
        reverse=True,
    )

    top_merchants = sorted(
        [{"merchant": name, "spending": amt} for name, amt in merchant_spend.items()],
        key=lambda row: row["spending"],
        reverse=True,
    )[:25]

    recurring_rows = []
    for (merchant, amount), dates in recurring_groups.items():
        unique_months = sorted({d.strftime("%Y-%m") for d in dates})
        if len(dates) < 3 or len(unique_months) < 3:
            continue
        sorted_dates = sorted(dates)
        gaps = [(sorted_dates[i] - sorted_dates[i - 1]).days for i in range(1, len(sorted_dates))]
        avg_gap = sum(gaps) / len(gaps) if gaps else 0
        if avg_gap <= 10:
            freq = "Weekly-ish"
        elif avg_gap <= 20:
            freq = "Fortnightly-ish"
        elif avg_gap <= 40:
            freq = "Monthly-ish"
        else:
            freq = "Irregular recurring"
        recurring_rows.append(
            {
                "merchant": merchant,
                "amount": amount,
                "count": len(dates),
                "months": ", ".join(unique_months),
                "frequency": freq,
            }
        )
    recurring_rows.sort(key=lambda row: (row["count"], row["amount"]), reverse=True)

    return {
        "income_total": income_total,
        "spending_total": spending_total,
        "monthly_rows": monthly_rows,
        "category_rows": category_rows,
        "top_merchants": top_merchants,
        "recurring_rows": recurring_rows,
    }


def add_table_headers(ws, row, headers):
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = TABLE_HEADER_FILL


def write_excel_report(output_path: Path, transactions: list[Transaction], analysis: dict):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    ws["A1"] = "Metric"
    ws["B1"] = "Value"
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    ws["A2"] = "Total Income"
    ws["B2"] = float(analysis["income_total"])
    ws["A3"] = "Total Spending"
    ws["B3"] = float(analysis["spending_total"])
    ws["A4"] = "Net Cashflow"
    ws["B4"] = float(analysis["income_total"] - analysis["spending_total"])
    for cell_ref in ("B2", "B3", "B4"):
        ws[cell_ref].number_format = "£#,##0.00"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 16

    ws_monthly = wb.create_sheet("Monthly Summary")
    add_table_headers(ws_monthly, 1, ["Month", "Income (£)", "Spending (£)", "Net (£)", "Running Net (£)"])
    for idx, row in enumerate(analysis["monthly_rows"], start=2):
        ws_monthly.cell(idx, 1, row["month"])
        ws_monthly.cell(idx, 2, float(row["income"]))
        ws_monthly.cell(idx, 3, float(row["spending"]))
        ws_monthly.cell(idx, 4, float(row["net"]))
        ws_monthly.cell(idx, 5, float(row["running_net"]))
    for col in ("B", "C", "D", "E"):
        for idx in range(2, len(analysis["monthly_rows"]) + 2):
            ws_monthly[f"{col}{idx}"].number_format = "£#,##0.00"

    category_rows = analysis.get("category_rows", [])
    month_rows = analysis.get("monthly_rows", [])

    for row in month_rows:
        month_label = row["month"].replace("-", " ")
        month_sheet = wb.create_sheet(month_label)
        add_table_headers(month_sheet, 1, ["Category", "Spending (£)"])
        month_totals = defaultdict(Decimal)
        for tx in sorted(transactions, key=lambda t: t.date):
            if tx.date.strftime("%Y-%m") != row["month"]:
                continue
            if tx.money_out <= 0 or tx.category == "":
                continue
            month_totals[tx.category] += tx.money_out
        for idx, (category, amount) in enumerate(sorted(month_totals.items(), key=lambda item: item[1], reverse=True), start=2):
            month_sheet.cell(idx, 1, category)
            month_sheet.cell(idx, 2, float(amount))
            month_sheet.cell(idx, 2).number_format = "£#,##0.00"

    ws_cat = wb.create_sheet("Category Breakdown")
    add_table_headers(ws_cat, 1, ["Category", "Spending (£)"])
    for idx, row in enumerate(category_rows, start=2):
        ws_cat.cell(idx, 1, row["category"])
        ws_cat.cell(idx, 2, float(row["spending"]))
        ws_cat.cell(idx, 2).number_format = "£#,##0.00"

    ws_tx = wb.create_sheet("Transactions")
    add_table_headers(
        ws_tx,
        1,
        [
            "Date",
            "Description",
            "Type",
            "Money In (£)",
            "Money Out (£)",
            "Balance (£)",
            "Category",
            "Category Reason",
            "Source Bank",
            "Source File",
        ],
    )
    for idx, tx in enumerate(sorted(transactions, key=lambda t: t.date), start=2):
        ws_tx.cell(idx, 1, tx.date.strftime("%d %b %Y"))
        ws_tx.cell(idx, 2, tx.description)
        ws_tx.cell(idx, 3, tx.tx_type)
        ws_tx.cell(idx, 4, float(tx.money_in))
        ws_tx.cell(idx, 5, float(tx.money_out))
        ws_tx.cell(idx, 6, float(tx.balance) if tx.balance is not None else None)
        ws_tx.cell(idx, 7, tx.category)
        ws_tx.cell(idx, 8, tx.category_reason)
        ws_tx.cell(idx, 9, tx.source_bank)
        ws_tx.cell(idx, 10, tx.source_file)
        for col in (4, 5, 6):
            ws_tx.cell(idx, col).number_format = "£#,##0.00"

    ws_verify = wb.create_sheet("Local Verification")
    verify_lines = local_verification_steps()
    for idx, line in enumerate(verify_lines, start=1):
        ws_verify.cell(idx, 1, line)
    ws_verify.column_dimensions["A"].width = 140

    ws_privacy = wb.create_sheet("Privacy Audit")
    add_table_headers(ws_privacy, 1, ["Library", "Type", "Network Connections?", "Notes"])
    for idx, (lib_name, lib_kind, network_flag, notes) in enumerate(PRIVACY_AUDIT, start=2):
        ws_privacy.cell(idx, 1, lib_name)
        ws_privacy.cell(idx, 2, lib_kind)
        ws_privacy.cell(idx, 3, network_flag)
        ws_privacy.cell(idx, 4, notes)
    ws_privacy.column_dimensions["A"].width = 20
    ws_privacy.column_dimensions["B"].width = 12
    ws_privacy.column_dimensions["C"].width = 24
    ws_privacy.column_dimensions["D"].width = 90

    wb.save(output_path)


def write_expense_breakdown_report(output_path: Path, transactions: list[Transaction], year: int, policy_path: Path):
    policy_data = {}
    if policy_path.exists():
        try:
            policy_data = json.loads(policy_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            policy_data = {}

    category_comments = policy_data.get("category_comments", {})
    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    def by_month_and_category(tx_list):
        month_totals = defaultdict(lambda: defaultdict(Decimal))
        for tx in tx_list:
            if tx.money_out <= 0 or not tx.category:
                continue
            month = tx.date.strftime("%b")
            month_totals[month][tx.category] += tx.money_out
        return month_totals

    ws_tx = Workbook()
    ws = ws_tx.active
    ws.title = "Transactions"
    headers = ["Date", "Month", "Description", "Type", "Money Out (£)", "Category", "Group", "Category Reason"]
    add_table_headers(ws, 1, headers)
    for idx, tx in enumerate(sorted(transactions, key=lambda t: t.date), start=2):
        if tx.money_out <= 0:
            continue
        ws.cell(idx, 1, tx.date.strftime("%d %b %Y"))
        ws.cell(idx, 2, tx.date.strftime("%b"))
        ws.cell(idx, 3, tx.description)
        ws.cell(idx, 4, tx.tx_type)
        ws.cell(idx, 5, float(tx.money_out))
        ws.cell(idx, 6, tx.category)
        ws.cell(idx, 7, "Essential" if tx.category in ESSENTIAL_CATEGORIES else "Discretionary")
        ws.cell(idx, 8, tx.category_reason)
        ws.cell(idx, 5).number_format = "£#,##0.00"

    month_totals = by_month_and_category(transactions)
    for month_index, month_name in enumerate(month_names, start=1):
        month_key = month_name
        ws_month = ws_tx.create_sheet(f"{month_name} {year}")
        ws_month["A1"] = "Essential"
        ws_month["B1"] = "2025"
        ws_month["C1"] = "Current"
        ws_month["D1"] = "In Retirement"
        ws_month["E1"] = "Comments"

        row_num = 2
        for category in ESSENTIAL_CATEGORIES:
            total = month_totals.get(month_key, {}).get(category, Decimal("0"))
            ws_month.cell(row_num, 1, category)
            ws_month.cell(row_num, 2, 0)
            ws_month.cell(row_num, 3, float(total))
            ws_month.cell(row_num, 4, 0)
            ws_month.cell(row_num, 5, category_comments.get(category, ""))
            ws_month.cell(row_num, 3).number_format = "£#,##0.00"
            row_num += 1

        ws_month.cell(row_num, 1, "TOTAL")
        ws_month.cell(row_num, 3, float(sum(month_totals.get(month_key, {}).get(cat, Decimal("0")) for cat in ESSENTIAL_CATEGORIES)))
        ws_month.cell(row_num, 3).number_format = "£#,##0.00"
        row_num += 2

        ws_month.cell(row_num, 1, "Discretionary")
        ws_month.cell(row_num, 2, "2025")
        ws_month.cell(row_num, 3, "Current")
        ws_month.cell(row_num, 4, "In Retirement")
        ws_month.cell(row_num, 5, "Comments")
        row_num += 1

        for category in DISCRETIONARY_CATEGORIES:
            total = month_totals.get(month_key, {}).get(category, Decimal("0"))
            ws_month.cell(row_num, 1, category)
            ws_month.cell(row_num, 2, 0)
            ws_month.cell(row_num, 3, float(total))
            ws_month.cell(row_num, 4, 0)
            ws_month.cell(row_num, 5, category_comments.get(category, ""))
            ws_month.cell(row_num, 3).number_format = "£#,##0.00"
            row_num += 1

        ws_month.cell(row_num, 1, "TOTAL")
        ws_month.cell(row_num, 3, float(sum(month_totals.get(month_key, {}).get(cat, Decimal("0")) for cat in DISCRETIONARY_CATEGORIES)))
        ws_month.cell(row_num, 3).number_format = "£#,##0.00"
        row_num += 2

        ws_month.cell(row_num, 1, "GRAND TOTAL")
        ws_month.cell(row_num, 3, float(sum(month_totals.get(month_key, {}).get(cat, Decimal("0")) for cat in ALL_CATEGORIES)))
        ws_month.cell(row_num, 3).number_format = "£#,##0.00"

    avg_ws = ws_tx.create_sheet(f"{year} Average")
    avg_ws["A1"] = "Essential"
    avg_ws["B1"] = "2025"
    avg_ws["C1"] = "Current"
    avg_ws["D1"] = "In Retirement"
    avg_ws["E1"] = "Comments"

    def avg_for_category(category: str):
        month_values = [month_totals.get(month_names[i], {}).get(category, Decimal("0")) for i in range(12)]
        if not month_values:
            return Decimal("0")
        return sum(month_values, Decimal("0")) / Decimal(len(month_values))

    row_num = 2
    for category in ESSENTIAL_CATEGORIES:
        avg = avg_for_category(category)
        avg_ws.cell(row_num, 1, category)
        avg_ws.cell(row_num, 2, 0)
        avg_ws.cell(row_num, 3, float(avg))
        avg_ws.cell(row_num, 4, 0)
        avg_ws.cell(row_num, 5, category_comments.get(category, ""))
        avg_ws.cell(row_num, 3).number_format = "£#,##0.00"
        row_num += 1
    avg_ws.cell(row_num, 1, "TOTAL")
    avg_ws.cell(row_num, 3, float(sum(avg_for_category(cat) for cat in ESSENTIAL_CATEGORIES)))
    avg_ws.cell(row_num, 3).number_format = "£#,##0.00"
    row_num += 2
    avg_ws.cell(row_num, 1, "Discretionary")
    row_num += 1
    for category in DISCRETIONARY_CATEGORIES:
        avg = avg_for_category(category)
        avg_ws.cell(row_num, 1, category)
        avg_ws.cell(row_num, 2, 0)
        avg_ws.cell(row_num, 3, float(avg))
        avg_ws.cell(row_num, 4, 0)
        avg_ws.cell(row_num, 5, category_comments.get(category, ""))
        avg_ws.cell(row_num, 3).number_format = "£#,##0.00"
        row_num += 1
    avg_ws.cell(row_num, 1, "TOTAL")
    avg_ws.cell(row_num, 3, float(sum(avg_for_category(cat) for cat in DISCRETIONARY_CATEGORIES)))
    avg_ws.cell(row_num, 3).number_format = "£#,##0.00"
    row_num += 2
    avg_ws.cell(row_num, 1, "GRAND TOTAL")
    avg_ws.cell(row_num, 3, float(sum(avg_for_category(cat) for cat in ALL_CATEGORIES)))
    avg_ws.cell(row_num, 3).number_format = "£#,##0.00"

    ws_tx.save(output_path)


def html_table(headers, rows):
    header_html = "".join(f"<th>{html.escape(str(h))}</th>" for h in headers)
    body_rows = []
    for row in rows:
        cols = "".join(f"<td>{html.escape(str(v))}</td>" for v in row)
        body_rows.append(f"<tr>{cols}</tr>")
    return f"<table><thead><tr>{header_html}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"


def local_verification_steps():
    return [
        "Local-only verification checklist:",
        "1) Inspect imports in this script: no requests/httpx/socket/urllib network clients are used.",
        "2) Search code for 'http', 'upload', 'telemetry', 'api', 'sync' to confirm no transmission logic.",
        "3) Run with network disabled; output generation should still complete.",
        "4) Monitor process network activity (e.g., TCPView/Resource Monitor) and verify zero outbound connections.",
        "5) Confirm outputs are local files only in the output directory.",
    ]


def privacy_audit_lines():
    lines = ["Privacy audit (normal operation):"]
    for lib_name, lib_kind, makes_network_connections, notes in PRIVACY_AUDIT:
        lines.append(
            f"- {lib_name} [{lib_kind}] | Network connections: {makes_network_connections} | {notes}"
        )
    return lines


def write_html_report(output_path: Path, analysis: dict):
    monthly_rows = [
        [
            row["month"],
            format_currency(row["income"]),
            format_currency(row["spending"]),
            format_currency(row["net"]),
            format_currency(row["running_net"]),
        ]
        for row in analysis["monthly_rows"]
    ]
    category_rows = [[row["category"], format_currency(row["spending"])] for row in analysis["category_rows"]]
    merchant_rows = [[row["merchant"], format_currency(row["spending"])] for row in analysis["top_merchants"]]
    recurring_rows = [
        [row["merchant"], format_currency(row["amount"]), row["count"], row["months"], row["frequency"]]
        for row in analysis["recurring_rows"]
    ]
    privacy_rows = [[lib_name, lib_kind, network_flag, notes] for lib_name, lib_kind, network_flag, notes in PRIVACY_AUDIT]

    html_content = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Local Spending Analysis Report</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; }}
    h1, h2 {{ color: #1f4e78; }}
    table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
    th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
    th {{ background: #d9e1f2; }}
    .metric {{ font-weight: bold; }}
    ul {{ line-height: 1.6; }}
  </style>
</head>
<body>
  <h1>Local Spending Analysis Report</h1>
  <p class="metric">Total income: {format_currency(analysis['income_total'])}</p>
  <p class="metric">Total spending: {format_currency(analysis['spending_total'])}</p>
  <p class="metric">Net cashflow: {format_currency(analysis['income_total'] - analysis['spending_total'])}</p>

  <h2>Monthly Summaries</h2>
  {html_table(["Month", "Income", "Spending", "Net", "Running Net"], monthly_rows)}

  <h2>Category Breakdown (Spending)</h2>
  {html_table(["Category", "Spending"], category_rows)}

  <h2>Top Merchants</h2>
  {html_table(["Merchant", "Spending"], merchant_rows)}

  <h2>Recurring Payments</h2>
  {html_table(["Merchant", "Amount", "Count", "Months", "Frequency"], recurring_rows)}

  <h2>Local-Only Verification</h2>
  <ul>
    {''.join(f"<li>{html.escape(step)}</li>" for step in local_verification_steps())}
  </ul>

  <h2>Privacy Audit (Libraries and Network Behavior)</h2>
  {html_table(["Library", "Type", "Network Connections During Normal Operation?", "Notes"], privacy_rows)}
</body>
</html>"""
    output_path.write_text(html_content, encoding="utf-8")


def write_privacy_audit_json(output_path: Path):
    payload = {
        "generated_at_local": datetime.now().isoformat(timespec="seconds"),
        "scope": "normal_operation",
        "libraries": [
            {
                "name": lib_name,
                "type": lib_kind,
                "network_connections_during_normal_operation": (network_flag == "Yes"),
                "notes": notes,
            }
            for lib_name, lib_kind, network_flag, notes in PRIVACY_AUDIT
        ],
    }
    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def export_transactions_csv(path: Path, transactions: list[Transaction]):
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "Date",
                "Description",
                "Type",
                "Money In (£)",
                "Money Out (£)",
                "Balance (£)",
                "Category",
                "Category Reason",
                "Source Bank",
                "Source File",
            ]
        )
        for tx in sorted(transactions, key=lambda t: t.date):
            writer.writerow(
                [
                    tx.date.strftime("%d %b %Y"),
                    tx.description,
                    tx.tx_type,
                    f"{tx.money_in:.2f}",
                    f"{tx.money_out:.2f}",
                    f"{tx.balance:.2f}" if tx.balance is not None else "",
                    tx.category,
                    tx.category_reason,
                    tx.source_bank,
                    tx.source_file,
                ]
            )


def should_skip_statement_file(path: Path) -> str | None:
    lower_name = path.name.lower()
    if any(marker in lower_name for marker in ("summary", "report", "tax_report", "allaccounts")):
        return "Skipped summary/export file"
    if any(marker in lower_name for marker in ("rainyday", "rainy day", "holidayspends", "holiday spends", "savings", "savings_")):
        return "Skipped savings/holiday account statements"
    return None


def read_transactions_from_folder(input_folder: Path, output_folder: Optional[Path], year: Optional[int]):
    transactions: list[Transaction] = []
    scanned_files = 0
    skipped_files = 0
    skip_reasons = defaultdict(int)
    output_root = output_folder.resolve() if output_folder is not None else None

    for path in input_folder.rglob("*"):
        if not path.is_file():
            continue

        resolved_path = path.resolve()
        if output_root is not None and resolved_path.is_relative_to(output_root):
            skipped_files += 1
            skip_reasons["Skipped generated output folder"] += 1
            continue

        skip_reason = should_skip_statement_file(path)
        if skip_reason:
            skipped_files += 1
            skip_reasons[skip_reason] += 1
            continue

        scanned_files += 1
        extension = path.suffix.lower()
        try:
            if extension == ".pdf":
                text = extract_pdf_text(path)
                provider = detect_pdf_provider(text)
                if provider == "lloyds":
                    rows = parse_lloyds_pdf(text, path.name)
                elif provider == "chase":
                    rows = parse_chase_pdf(text, path.name)
                elif provider == "capital_one":
                    rows = parse_capital_one_pdf(text, path.name)
                else:
                    skipped_files += 1
                    skip_reasons["Unsupported PDF layout"] += 1
                    rows = []
                transactions.extend(rows)
            elif extension == ".csv":
                parent_name = path.parent.name.lower()
                if "statement" in parent_name or any(marker in path.name.lower() for marker in ("statement", "transaction")):
                    transactions.extend(parse_csv_file(path))
                else:
                    skipped_files += 1
                    skip_reasons["Skipped non-statement CSV export"] += 1
            else:
                skipped_files += 1
                skip_reasons["Unsupported file extension"] += 1
        except (PdfReadError, OSError, ValueError):
            skipped_files += 1
            skip_reasons["Unreadable or invalid file"] += 1
            continue

    if year is not None:
        transactions = [tx for tx in transactions if tx.date.year == year]
    return transactions, scanned_files, skipped_files, dict(skip_reasons)


def main():
    parser = argparse.ArgumentParser(
        description="Local-only bank statement parser and spending analysis tool."
    )
    parser.add_argument(
        "--input-folder",
        required=True,
        help="Local folder containing statement PDFs/CSVs (searched recursively).",
    )
    parser.add_argument(
        "--output-folder",
        default=DEFAULT_OUTPUT_FOLDER,
        help=(
            "Local output folder for Excel/HTML/CSV files "
            "(default: G:\\My Drive\\Statements\\Analysis Outputs\\Local Statement Tool)."
        ),
    )
    parser.add_argument(
        "--year",
        type=int,
        help="Optional year filter, e.g., 2026.",
    )
    parser.add_argument(
        "--policy-json",
        default=str(DEFAULT_POLICY_JSON),
        help="Optional local JSON policy file for deterministic hint-based categorization.",
    )
    parser.add_argument(
        "--excel-name",
        default="local_spending_analysis.xlsx",
        help="Output Excel filename.",
    )
    parser.add_argument(
        "--html-name",
        default="local_spending_analysis.html",
        help="Output HTML filename.",
    )
    parser.add_argument(
        "--transactions-csv-name",
        default="normalized_transactions_local.csv",
        help="Output normalized transactions CSV filename.",
    )
    parser.add_argument(
        "--privacy-audit-json",
        default="privacy_audit_local.json",
        help="Output privacy audit JSON filename.",
    )
    args = parser.parse_args()

    input_folder = Path(args.input_folder)
    output_folder = Path(args.output_folder)
    policy_path = Path(args.policy_json)
    if not input_folder.exists():
        raise FileNotFoundError(f"Input folder not found: {input_folder}")
    output_folder.mkdir(parents=True, exist_ok=True)

    normalize_patterns, policy_hints = load_policy(policy_path)
    transactions, scanned_files, skipped_files, skip_reasons = read_transactions_from_folder(
        input_folder,
        output_folder,
        args.year,
    )
    analysis = analyse_transactions(transactions, normalize_patterns, policy_hints)

    year_name = str(args.year) if args.year is not None else "all"
    year_filter = args.year

    excel_path = output_folder / args.excel_name
    html_path = output_folder / args.html_name
    transactions_csv_path = output_folder / args.transactions_csv_name
    privacy_audit_json_path = output_folder / args.privacy_audit_json

    write_excel_report(excel_path, transactions, analysis)
    expense_breakdown_path = output_folder / f"expense_breakdown_{year_name}.xlsx"
    write_expense_breakdown_report(expense_breakdown_path, transactions, year_filter or datetime.now().year, policy_path)
    write_html_report(html_path, analysis)
    export_transactions_csv(transactions_csv_path, transactions)
    write_privacy_audit_json(privacy_audit_json_path)

    print("Local-only processing complete.")
    print(f"Input folder: {input_folder.resolve()}")
    print(f"Files scanned: {scanned_files}")
    print(f"Files skipped: {skipped_files}")
    if skip_reasons:
        print(f"Skip reasons: {skip_reasons}")
    print(f"Transactions parsed: {len(transactions)}")
    print(
        "This summary is derived from local statement contents read during this run; "
        "no statement data was transmitted externally."
    )
    print(f"Total income: {format_currency(analysis['income_total'])}")
    print(f"Total spending: {format_currency(analysis['spending_total'])}")
    print(f"Excel report: {excel_path.resolve()}")
    print(f"HTML report: {html_path.resolve()}")
    print(f"Normalized CSV: {transactions_csv_path.resolve()}")
    print(f"Privacy audit JSON: {privacy_audit_json_path.resolve()}")
    print("Dependencies: pypdf, openpyxl (plus Python standard library).")
    print("\n".join(privacy_audit_lines()))
    print("\n".join(local_verification_steps()))


if __name__ == "__main__":
    main()
