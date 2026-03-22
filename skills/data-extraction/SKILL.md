---
name: Data Extraction
description: >
  Activates when the user asks to "extract data", "run extract", "build Extract Balance",
  "build Extract GL", "parse livre de paie", "parse charges patronales", "créer les feuilles
  Extract", "extraire les données", "populate extract sheets", or when any extraction step
  in the audit workflow begins. Guides parsing each source file with the correct parser script
  and writing the four Extract sheets into the FT-P-2 workbook.
version: 1.1.0
---

## Data Extraction — Skill Guide (v1.1 — SYSCOHADA corrections)

This skill governs the creation of all four Extract sheets inside the FT-P-2 workbook.

---

### Step 0 — Accounting Plan Check (MANDATORY FIRST STEP)

Before any extraction, verify `.audit-session.json` contains `"accounting_plan"`.

If absent, use **AskUserQuestion**:
> "Quel référentiel comptable utilise cette société? / Which accounting standard?"
> - SYSCOHADA (Cameroun, zone OHADA) → comptes **66x** pour les charges de personnel
> - PCG / France → comptes **64x** pour les charges de personnel

Store result as `"accounting_plan": "SYSCOHADA"` or `"PCG"` in the session.

**Account ranges by plan:**

| Plan | Rémunérations directes | Cotisations sociales | CF/P+FNE | Autres info |
|------|----------------------|---------------------|----------|-------------|
| SYSCOHADA | 661x + 663x | 664110/120/130 | 664380 + FNE=0 | 668x |
| PCG/France | 641x + 642x | 645xxx | 647xxx | 648x |

---

### Step 1 — Filter Transparency (MANDATORY before any filter)

**Before applying any filter** (journal code, account range, CSV rubrique), always:

1. Compute the impact table from the raw data:

```
| Critère     | Valeur retenue | Montant Débit    | % du total | Inclus |
|-------------|----------------|------------------|------------|--------|
| Journal PAY | exclu          | 651 500 000 FCFA |    97.1 %  | Non    |
| Journal CAM | retenu         |  19 216 076 FCFA |     2.9 %  | Oui    |
```

2. Show the table to the user.
3. Use **AskUserQuestion** to confirm before executing.
4. If excluded amount > 5% of total: add an explicit warning.

---

### Four Extract Sheets — TCD/Pivot Format

**All Extract sheets must use pivot format (one row per employee), NOT raw source rows.**

---

#### A — "Extract LivrePaie"

Source: CSV Livre de Paie

Pivot by employee, rubrique `BRUT` only:

| Etiquette de lignes | NOM | PRENOM | Salaire BRUT |
|---------------------|-----|--------|-------------|
| 001234 | DUPONT | Jean | 450 000 |
| ... | | | |
| **TOTAL** | | | **sum** |

Rules:
- Matricule filter: keep only rows where Matricule matches `^\d{3,}` — exclude rows where Matricule = "Total", "TOTAL", or any non-numeric string.
- Code filter: exact match `TypeCode == "BRUT"` (case-sensitive, no substring match).
- One row per Matricule (outer key), aggregated with sum.
- Last row: label = "TOTAL", NOM/PRENOM blank, amount = grand total.
- Sort by Matricule ascending.

---

#### B — "Extract Charges Patronal"

Source: CSV Charges Patronales

Pivot by employee, 5 patronal charge codes:

| Etiquette de lignes | NOM | PRENOM | Crédit Foncier Patronal | Fond National de l'emploi (FNE) | Pension Vieillesse (CNPS) | Allocation Familiale | Accident de Travail |
|---|---|---|---|---|---|---|---|
| 001234 | ... | | cf/p | fne | cnps | af | at |
| ... | | | | | | | |
| Total | = | = | ΣCF/P | ΣFNE | ΣCNPS | ΣAF | ΣAT |
| **TOTAL** | | | 2×ΣCF/P | 2×ΣFNE | 2×ΣCNPS | 2×ΣAF | 2×ΣAT |

Rules:
- Matricule filter: `^\d{3,}` (same as LivrePaie — exclude "Total" rows).
- Code mapping (exact TypeCode):
  - `4100` → Crédit Foncier Patronal (CF/P)
  - `4400` → Fond National de l'Emploi (FNE)
  - `4500` → Pension Vieillesse (CNPS/P)
  - `4800` → Allocation Familiale (AF)
  - `4900` → Accident de Travail (AT)
- Penultimate row: label = "Total", NOM = "=", PRENOM = "=", amounts = column sums.
- Last row: label = "TOTAL", amounts = 2 × column sums (patronal + salarial parts equal for CF/P+FNE in SYSCOHADA Cameroun).

---

#### C — "Extract GL"

Source: Grand Livre Général (xls)

Raw filtered rows (not pivoted) — **compte 661800 + journal CAM only**:

Columns: `Compte | Date | Code Journal | N° Pièce | Libellé | Libellé 2 | Débit | Crédit | Solde Cumulé`

Rules:
- Filter: `Compte == "661800"` AND `CodeJournal == "CAM"`
- Solde Cumulé = running cumulative balance (recalculated line by line: Débit − Crédit cumulated from top)
- Write all filtered rows as-is, no aggregation.

**Filter transparency required:** Before applying the journal/account filter, show a breakdown of all journals + all accounts with their debit amounts and ask confirmation.

---

#### D — "Extract Balance"

Source: Balance Générale (xlsx) — **3-part structured format with subtotals**

Column structure: `Compte | Libellé | Mvt Débit | Mvt Crédit | Solde Net`

Where `Solde Net = Mvt Débit − Mvt Crédit`

**Part 1 — Rémunérations directes et avantages (661x + 663x):**
- All accounts starting with `661` or `663` with non-zero Solde Net
- Subtotal row label: `"Sous-total 661-663"` → feeds Section A of Feuil2 (column R = SAL BRUT)

**Part 2 — Cotisations CNPS (664110, 664120, 664130):**
- Exact accounts: 664110 (AF), 664120 (CNPS/P Pension AV), 664130 (AT)
- Subtotal row label: `"Sous-total CNPS"` → feeds Section B of Feuil2 (columns S, W, X)

**Part 3 — CF/P et FNE (664380 + FNE note):**
- Exact account: 664380 (Provision CF/P)
- FNE note row: label = "FNE — non comptabilisé en GL", Solde Net = 0
- Subtotal row label: `"Sous-total CF/P+FNE"` → feeds Section C of Feuil2 (columns T, U)

**Note:** 668x accounts do NOT appear in Extract Balance — they are informative only (Section D of Feuil2, fed directly from GL).

---

### CSV Parsing — Critical Rules

**Format:** Both CSV files use Excel text-forcing `="value"` with `;` delimiter, latin-1 encoding.

**Do NOT use `pandas.read_csv()` directly.** Use the line-by-line parser:
```python
import re
def parse_csv_value(raw):
    s = raw.strip()
    s = re.sub(r'^=""?', '', s)
    s = re.sub(r'""?$', '', s)
    return s

def parse_amount(raw):
    s = parse_csv_value(raw)
    s = s.replace('","', '.').replace(',', '.').replace(' ', '')
    try:
        return float(s)
    except ValueError:
        return 0.0
```

**Matricule filtering rule:** `re.match(r'^\d{3,}', matricule)` — must start with at least 3 digits. This excludes "Total", "TOTAL", empty rows, and header rows.

**BRUT code:** TypeCode exact match `== "BRUT"` — never substring, never case-insensitive.

---

### BG Column Positions (Balance Générale xlsx)

Standard 8-column format:
- col 0 = Compte (account code, 6 digits)
- col 1 = Libellé
- col 2 = Solde Débit Ouverture
- col 3 = Solde Crédit Ouverture
- col 4 = **Mouvement Débit** ← use this
- col 5 = **Mouvement Crédit** ← use this
- col 6 = Solde Débit Clôture
- col 7 = Solde Crédit Clôture

Solde Net = col4 − col5

---

### Error Handling

- If CSV has 0 employees after Matricule filter: warn that the filter may be incorrect, show 5 sample Matricule values, and ask user to confirm or adjust.
- If BG file has fewer than 8 columns: warn and ask user to verify the file format.
- After each extraction: print `✅ Extract [Name]: N rows extracted`. If 0 rows: print `⚠️ 0 rows — verify source file and filters`.
