---
name: Data Extraction
description: >
  Activates when the user asks to "extract data", "run extract", "build Extract Balance",
  "build Extract GL", "parse livre de paie", "parse charges patronales", "créer les feuilles Extract",
  "extraire les données", "populate extract sheets", or when any extraction step in the audit
  workflow begins. Guides parsing each source file with the correct parser script and writing
  the four Extract sheets into the FT-P-2 workbook.
version: 1.0.0
---

## Data Extraction — Skill Guide

This skill governs the creation of all four Extract sheets inside the FT-P-2 workbook.
Run each parser script in order; each one writes directly into the workbook.

---

### Four Extract Sheets to Build

| Sheet name | Source file | Parser script | Key filter/transform |
|------------|-------------|---------------|---------------------|
| Extract Balance | Balance Générale (.xlsx) | `parse_balance.py` | Rows where Rubrique/Libellé contains "Charges du personnel" OR account 661–663; NetSolde = MvtDebit − MvtCredit |
| Extract GL | Grand Livre Général (.xls) | `parse_grand_livre.py` | CodeJournal == 'CAM' AND account starts with '66' |
| Extract Charges Patronal | Charges Patronales (.CSV) | `parse_charges_patronales.py` | Pivot: one row per employee; columns = CF/P, FNE, CNPS/P (Pension Vieillesse), AF (Allocation Familiale), AT (Accident de Travail) |
| Extract LivrePaie | Livre de Paie (.CSV) | `parse_livre_paie.py` | Pivot: one row per employee; column = SAL BRUT |

---

### Balance Générale Parsing

Structure (8 columns): `Compte | Libellé | SolDebitOuv | SolCreditOuv | MvtDebit | MvtCredit | SolDebitClo | SolCreditClo`

NetSolde formula: **`MvtDebit − MvtCredit`** (for expense accounts in class 66, debit = positive charge)

Filter: keep rows where account starts with `661`, `662`, or `663`.

---

### Grand Livre Parsing

File format: `.xls` (legacy Excel) — requires `xlrd`. Sheet name: `Sage`.

Columns (11): `Compte | Date | CodeJournal | NoPiece | Libelle | Libelle2 | _d2 | _flag | Debit | Credit | Solde`

Filter: `CodeJournal == 'CAM'` AND `Compte` starts with `'66'`

---

### CSV File Parsing (Charges Patronales & Livre de Paie)

Both CSV files use a non-standard Excel text-forcing format:
- Delimiter: `;`
- Encoding: `latin-1`
- Values wrapped in `="value"` or `=""value""`

**Critical parsing pattern** — do NOT use `pandas.read_csv` directly; use line-by-line parsing:
```python
import re
def parse_csv_value(raw):
    cleaned = re.sub(r'^=""?', '', raw.strip('"'))
    cleaned = re.sub(r'""?$', '', cleaned)
    return cleaned
```

Amount fields use French decimal comma: `87453","00` → after parsing → `87453.00`
Amount fix: `cleaned.replace('","', '.').replace(',', '.')`

**Charges Patronales codes** (column index 4 = TypeCode):
- `4100` = Crédit Foncier Patronal (CF/P) → column T in Feuil2
- `4400` = Fond National de l'Emploi (FNE) → column U in Feuil2
- `4500` = CNPS Pension Vieillesse → column S in Feuil2
- `4800` = Allocation Familiale (AF) → column W in Feuil2
- `4900` = Accident de Travail (AT) → column X in Feuil2

**Livre de Paie key code**: `BRUT` = Salaire Brut → column R in Feuil2

---

### Pivot Table Structure

Both CSV parsers produce a pivot with employee Matricule as the row key, merged with Nom/Prénom.
Use `pandas.pivot_table(aggfunc='sum')` then `reset_index()`.

The Matricule field is the join key between LivrePaie (for BRUT) and ChargesPatronales (for S/T/U/W/X).

---

### Sheet Formatting Standards

Each Extract sheet must:
- Have a bold header row (Arial 10pt, `1F4E79` blue fill, white text)
- Have alternating row fills (`F5F8FC` / white)
- Apply `#,##0;(#,##0);"-"` number format to all numeric columns
- Have column widths auto-fitted (min 10, max 40 chars)
- Have auto-filter enabled on the header row
- Have freeze_panes on the row below the header

---

### Error Handling

- If a source file is not found, read path from `.audit-session.json`; if absent, use AskUserQuestion.
- If the Grand Livre has no rows matching the filter, warn the user (may indicate different journal code).
- If a CSV employee appears in ChargesPatronales but not in LivrePaie (or vice versa), include them
  with NaN → 0 for the missing columns and log a warning.
- Always print row counts after each extraction: `Extracted N rows → sheet 'Extract Balance'`
