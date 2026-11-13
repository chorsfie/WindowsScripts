# Local Statement Tool User Guide

This guide covers the local-only statement analysis tool in:

- [_Scripts/local_statement_tool.py](_Scripts/local_statement_tool.py)

## Purpose

The tool parses bank statements from local folders and produces spending analysis reports without sending data externally.

## Privacy and Local-Only Operation

The tool is designed to run fully locally:

- No cloud APIs
- No AI/LLM APIs
- No telemetry/upload/sync behavior
- Reads statement files from local disk only
- Writes outputs to local disk only

Privacy audit output is generated each run as JSON.

## Supported Inputs

- PDF statements:
  - Lloyds
  - Chase
  - Capital One
- CSV statement exports (generic column matching)

## Required Dependencies

- `pypdf`
- `openpyxl`

Install:

```powershell
py -m pip install pypdf openpyxl
```

## Recommended Folder Layout

Statement sources (examples):

- `G:\My Drive\Statements\Lloyds Statements`
- `G:\My Drive\Statements\Chase Statements`
- `G:\My Drive\Statements\Capital One Statements`

Recommended output folder:

- `G:\My Drive\Statements\Analysis Outputs\Local Statement Tool`

Important: the tool automatically skips its own generated output folder when it is nested under the input folder, so it does not re-read Excel/HTML/CSV outputs created during the same run. This prevents double-counting and date distortions caused by ingesting generated report files.

## Run Command

```powershell
$repoRoot = 'C:\path\to\your\repo'

py "$repoRoot\_Scripts\local_statement_tool.py" `
  --input-folder 'G:\My Drive\Statements' `
  --output-folder 'G:\My Drive\Statements\Analysis Outputs\Local Statement Tool' `
  --year 2026 `
  --policy-json "$repoRoot\_Scripts\expense_policy_hints.json" `
  --excel-name 'local_spending_analysis_2026.xlsx' `
  --html-name 'local_spending_analysis_2026.html' `
  --transactions-csv-name 'normalized_transactions_local_2026.csv' `
  --privacy-audit-json 'privacy_audit_local_2026.json'
```

If `--output-folder` is omitted, the tool now defaults to:

- `G:\My Drive\Statements\Analysis Outputs\Local Statement Tool`

## Outputs

The tool generates:

- Excel report (`.xlsx`)
- HTML report (`.html`)
- Normalized transactions CSV
- Privacy audit JSON

## Deterministic Categorization

Categorization uses deterministic rules only:

- Built-in keyword patterns in script
- Optional policy hints in:
  - [_Scripts/expense_policy_hints.json](_Scripts/expense_policy_hints.json)

No probabilistic/AI classification is used.

## Verification Checklist (Local-Only)

1. Review imports in the script (no network client libraries used by tool logic).
2. Search script for `upload`, `telemetry`, `sync`, `api`.
3. Run with internet disconnected and confirm outputs still generate.
4. Monitor process network activity (Resource Monitor/TCPView) during execution.
5. Confirm all outputs are written to local folder paths.
