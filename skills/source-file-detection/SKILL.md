---
name: Source File Detection
description: >
  Activates when the user starts a payroll audit, runs /detect-files, asks "which files do I need",
  mentions "find source files", "scan for audit files", "fichiers source", "détection des fichiers",
  "where are the payroll files", "locate balance générale", or initiates the audit workflow.
  Provides step-by-step guidance for scanning the working directory, identifying the required source
  files (Balance Générale, Grand Livre, Livre de Paie, Charges Patronales), and confirming or
  requesting file paths from the user before any extraction begins.
version: 1.0.0
---

## Source File Detection — Skill Guide

When this skill activates, scan the user's working directory for required audit source files and
confirm matches with the user before proceeding. Always use AskUserQuestion for confirmation steps.

---

### Required Source Files

Every payroll audit requires exactly these four source files:

| File type | Expected filename pattern | Description |
|-----------|--------------------------|-------------|
| Balance Générale | `BG*.xlsx`, `Balance*.xlsx`, `BALANCE*.xlsx` | Trial balance; accounts 661–663 = Charges du personnel |
| Grand Livre Général | `Grand Livre*.xls*`, `GL*.xls*`, `GRAND LIVRE*.xls*` | General ledger; filter journal `CAM` + accounts `66` |
| Livre de Paie | `LIVREPAIE*.CSV`, `Livre*Paie*.csv`, `LivrePaie*.csv` | Payroll register; contains SAL BRUT per employee |
| Charges Patronales | `CHARGESPATRON*.CSV`, `Charges*Patron*.csv` | Employer charges; contains CF/P, FNE, CNPS, AF, AT per employee |

And one output workbook (template):

| File type | Expected pattern | Description |
|-----------|-----------------|-------------|
| Feuille de Travail | `FT-P-2*.xlsx` | Output workbook; all extracts and reconciliation go here |

---

### Detection Algorithm

1. **Run `scripts/detect_files.py`** in the user's working directory (the directory where they
   ran Claude Code, or a path they specify).

2. **Present findings** using a table: ✅ Found / ⚠️ Multiple matches / ❌ Not found.

3. **Use AskUserQuestion** for each file that is:
   - Not found (ask user to provide path)
   - Has multiple matches (ask user to confirm which one)

4. **Template special case**: If no `FT-P-2*.xlsx` exists, copy `assets/FT-P-2-template.xlsx`
   to the working directory and ask the user what name to give the output file.
   Default name: `FT-P-2- Exhaustivité des charges du personnel (Salaire brut).xlsx`

5. **Persist confirmed paths** to `.audit-session.json` in the working directory so other
   agents/commands can read them without re-asking.

---

### Session State File (.audit-session.json)

Every confirmed file path is stored in:
```json
{
  "language": "fr",
  "run_mode": "interactive",
  "working_dir": "/path/to/audit/folder",
  "files": {
    "balance_generale": "/path/to/BG CIFM 2025.xlsx",
    "grand_livre": "/path/to/GRAND LIVRE GENERALE CIFM 2025.xls",
    "livre_paie": "/path/to/LIVREPAIE.CSV",
    "charges_patronales": "/path/to/CHARGESPATRONALES.CSV",
    "feuille_travail": "/path/to/FT-P-2- Exhaustivité....xlsx"
  },
  "steps_completed": []
}
```

Always write this file after confirmation. Always read it at the start of any subsequent step.

---

### Language Handling

- Ask the user their preferred language at the start of every new session if `.audit-session.json`
  does not yet exist.
- Store choice in `language` field: `"fr"` or `"en"`.
- Use bilingual labels in AskUserQuestion options: `"Balance Générale / Trial Balance"`.
- Print confirmation messages in the user's chosen language.

---

### Error Handling

- If the working directory has no `.xlsx`/`.xls`/`.CSV` files at all, warn the user clearly
  and ask them to navigate to the correct folder first (`cd /path/to/audit`).
- If a CSV file is found but has `.csv` (lowercase), still match it (case-insensitive glob).
- If the Grand Livre is an `.xls` (old Excel format), note that `xlrd` is required.

---

### Package Pre-Check

Before any script runs, execute `scripts/check_packages.py` which verifies:
- `pandas`, `openpyxl`, `xlrd`, `numpy` are installed.
- If any are missing, print the exact `pip install` command and stop until the user confirms
  packages are installed.

---

### Reference: detect_files.py

The script `scripts/detect_files.py` performs all glob matching and writes `.audit-session.json`.
It accepts an optional `--dir` argument to scan a specific path.
It outputs a JSON summary of found/missing files to stdout for the agent to parse.
