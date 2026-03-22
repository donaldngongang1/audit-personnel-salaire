# audit-personnel-salaire

**Automated payroll audit workpaper plugin for Claude Code**

Builds a complete reconciliation between payroll records (Livre de Paie) and accounting entries
(Balance Générale / Grand Livre) for OHADA/Cameroonian chart of accounts — class 66 (Charges du personnel).

Generates:
- 4 Extract sheets (Balance Générale, Grand Livre, Charges Patronales, Livre de Paie)
- Complete Feuil2 reconciliation workpaper (PAIE vs COMPTABILITÉ with ECART row)
- Feuil1 executive summary with automated gap explanations
- Continuous verification via 3 independent eval scripts

## Features

- **Bilingual (FR/EN)** — All prompts, labels, and explanations in French and English
- **Interactive or Unattended** — User chooses at startup
- **Auto-detection** — Scans working directory for source files; asks for paths if not found
- **Template management** — Copies `FT-P-2-template.xlsx` to the user's directory; asks for filename
- **Package check** — Verifies required Python packages before running any script
- **Continuous verification** — Eval scripts run after every write to catch errors immediately
- **Gap analysis** — Automatically explains known structural gaps (FNE, timing differences, rounding)

## Required Source Files

| File | Description |
|------|-------------|
| `BG*.xlsx` | Balance Générale (Trial Balance) |
| `Grand Livre*.xls` | Grand Livre Général (General Ledger) |
| `LIVREPAIE*.CSV` | Livre de Paie (Payroll Register) |
| `CHARGESPATRON*.CSV` | Charges Patronales (Employer Charges) |

## Required Python Packages

```bash
pip install pandas openpyxl xlrd numpy
```

## Usage

### Quick Start
```
/audit-personnel-salaire:run
```

### Step by Step
```
/audit-personnel-salaire:detect-files
/audit-personnel-salaire:extract
/audit-personnel-salaire:reconcile
/audit-personnel-salaire:verify
/audit-personnel-salaire:summarise
```

## Account Mapping (OHADA Class 66)

| Column | Account(s) | Description |
|--------|-----------|-------------|
| R — SAL BRUT | 661xxx, 663xxx | Salaires, appointements, allocations |
| S — CNPS/P | 664120 | CNPS Pension Vieillesse (AV) |
| T — CF/P | 664380 | Crédit Foncier Patronal / Provision |
| U — FNE | — | FNE (structural 0 in GL) |
| W — AF | 664110 | CNPS Allocation Familiale |
| X — AT | 664130 | CNPS Accident de Travail |

## Installation

Install from GitHub via Claude Code:
```
cc plugin install https://github.com/XavierNGONGANGACHOUN/audit-personnel-salaire
```

Or locally for development:
```
cc --plugin-dir /path/to/audit-personnel-salaire
```

## License

Commercial. All rights reserved — Xavier Ngonganga Choun.
