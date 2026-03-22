---
name: extract
description: "Run data extraction: parse Balance Générale, Grand Livre, Charges Patronales, and Livre de Paie into the four Extract sheets of the FT-P-2 workbook."
argument-hint: "[--sheet balance|gl|paie|charges|all]"
allowed-tools: Bash, Read, AskUserQuestion
---

## /audit-personnel-salaire:extract

Parse source files and populate Extract sheets in the FT-P-2 workbook.

### Pre-check

Read `.audit-session.json`. If it doesn't exist, run detect-files first:
"Session not found. Please run /detect-files first. / Session introuvable. Veuillez d'abord exécuter /detect-files."

Check packages: `python scripts/check_packages.py`

### Extraction by sheet

Determine which sheet(s) to extract based on `--sheet` argument (default = all):

```bash
# All extractions
python scripts/build_extracts.py --sheet all

# Or individually:
python scripts/parse_balance.py          # → Extract Balance
python scripts/parse_grand_livre.py      # → Extract GL
python scripts/parse_charges_patronales.py  # → Extract Charges Patronal
python scripts/parse_livre_paie.py       # → Extract LivrePaie
```

### After each extraction

Show: `✅ Extract [Sheet Name]: N rows extracted → [FT-P-2 filename]`

If row count is 0 for a required sheet, warn:
`⚠️ Extract [Sheet Name]: 0 rows — verify source file and filter criteria.`

Use AskUserQuestion if extraction fails to ask whether to retry with a different file path.

### Completion

Print summary of all Extract sheets with row counts.
Update `steps_completed` in `.audit-session.json`.
