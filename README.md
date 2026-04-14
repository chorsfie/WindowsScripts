# Craig's GIT Workspace

This workspace contains a mix of Python and PowerShell utilities for document search, Microsoft Entra policy export, and bank statement parsing.

## Workspace layout

| Folder | Purpose |
|---|---|
| `EpsteIn/` | Python CLI tool that searches indexed Epstein files for mentions of LinkedIn connections and generates an HTML report. |
| `WindowsScripts/` | PowerShell scripts for parsing bank statement PDFs into CSV/summary outputs. |
| `_Scripts/` | PowerShell scripts and outputs for Microsoft Entra / Conditional Access exports. |
| `project_wenv/` | Shared Python virtual environment at workspace root. |

---

## EpsteIn (Python)

- Main script: `EpsteIn/EpsteIn.py`
- Docs: `EpsteIn/README.md`
- Dependencies: `EpsteIn/requirements.txt`

### Quick run (Windows PowerShell)

```powershell
cd EpsteIn
..\project_wenv\Scripts\Activate.ps1
pip install -r requirements.txt
python .\EpsteIn.py --connections C:\path\to\Connections.csv
```

Output defaults to `EpsteIn/EpsteIn.html`.

---

## WindowsScripts (PowerShell)

- `WindowsScripts/BankStatementParser.ps1`
  - Generic parser for UK bank PDFs (Chase/Lloyds detection)
  - Produces transaction exports and summary outputs
- `WindowsScripts/ChasePDFParsing.ps1`
  - Earlier Chase-focused parser variant

### Prerequisite

Both scripts require Poppler `pdftotext` and a valid path configured in the script variable:

```powershell
$popplerPath = "C:\Program Files\Poppler\...\pdftotext.exe"
```

### Quick run

```powershell
cd WindowsScripts
.\BankStatementParser.ps1
```

or

```powershell
cd WindowsScripts
.\ChasePDFParsing.ps1
```

---

## _Scripts (PowerShell / Entra)

- `_Scripts/Export-ConditionalAccessPolicies.ps1`
  - Connects to Microsoft Graph
  - Resolves users/groups/roles in Conditional Access policy assignments
  - Exports CSV output

- `_Scripts/ConditionalAccessPolicies.csv`
  - Export output file (currently empty in this workspace snapshot)

- `_Scripts/EntraAdmin/`
  - Reserved folder (currently empty)

### Quick run

```powershell
cd _Scripts
.\Export-ConditionalAccessPolicies.ps1
```

The script prompts for Microsoft Graph sign-in and writes `ConditionalAccessPolicies.csv`.

---

## Notes

- This repo mixes executable scripts and generated outputs in the same folders.
- If you want cleaner version control, consider adding generated CSV/report files to `.gitignore`.
- For project-specific details, prefer each folder's own README (for example `EpsteIn/README.md`).
