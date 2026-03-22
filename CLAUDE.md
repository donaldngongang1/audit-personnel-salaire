# audit-personnel-salaire — Plugin Rules & Regression Reference

This file documents the 6 mandatory business rules for the payroll audit plugin,
derived from a real audit session (client CIFM, exercice 2025).

---

## RULE 1 — Accounting Plan: SYSCOHADA ≠ PCG France

**Always ask the accounting plan at the start of each audit session.**

| Plan | Charges de personnel | Cotisations | CF/P+FNE | Autres |
|------|---------------------|------------|---------|-------|
| SYSCOHADA (Cameroun/OHADA) | **661x + 663x** | **664110/120/130** | **664380 + FNE=0** | 668x (info) |
| PCG France | 641x + 642x | 645xxx | 647xxx | 648x |

Store as `.audit-session.json` → `"accounting_plan": "SYSCOHADA"` or `"PCG"`.

If not specified: ask via AskUserQuestion before any extraction.

---

## RULE 2 — Filter Transparency

**Before applying any filter (journal code, account range, CSV rubrique):**

1. Compute and display an impact table with amounts and percentages:
```
| Journal | Montant Débit    | % du total | Inclus |
|---------|------------------|------------|--------|
| PAY     | 651 500 000 FCFA |    97.1 %  | Non    |
| CAM     |  19 216 076 FCFA |     2.9 %  | Oui    |
```
2. Ask confirmation via AskUserQuestion.
3. If excluded > 5% of total: add explicit warning before proceeding.

---

## RULE 3 — Extract Sheet Format: TCD/Pivot (Not Raw Rows)

**All Extract sheets must use pivot format — one row per employee, not raw source rows.**

### Extract LivrePaie
- Rubrique BRUT only (exact code match `== "BRUT"`)
- Matricule filter: `^\d{3,}` (exclude "Total", "TOTAL", non-numeric)
- Last row: `TOTAL` with grand sum

### Extract Charges Patronal
- Codes: 4100=CF/P | 4400=FNE | 4500=CNPS/P | 4800=AF | 4900=AT
- Matricule filter: `^\d{3,}`
- Penultimate row: `Total` with `=` in NOM/PRENOM, column sums
- Last row: `TOTAL` with 2× column sums (patronal + salarial parts equal in Cameroun)

### Extract GL
- Raw rows: compte 661800 + journal CAM only
- Solde Cumulé = running cumulative balance

### Extract Balance
- 3 parts from BG (col4=MvtDebit, col5=MvtCredit), each with subtotal:
  - Part 1: 661x+663x → subtotal feeds Feuil2 col R (SAL BRUT)
  - Part 2: 664110/120/130 → subtotal feeds Feuil2 cols S/W/X (CNPS)
  - Part 3: 664380 + FNE note (=0) → subtotal feeds Feuil2 cols T/U (CF/P+FNE)

---

## RULE 4 — Source of Truth by Section

| Section Feuil2 | Comptes | Source | Reason |
|----------------|---------|--------|--------|
| A — Rémunérations | 661x + 663x | **Balance Générale (BG xlsx)** | BG exhaustive; GL may miss entries |
| B — CNPS | 664110/120/130 | **Grand Livre (GL xls), tous journaux** | GL gives per-account breakdown |
| C — CF/P+FNE | 664380 + FNE=0 | **Grand Livre (GL xls), tous journaux** | FNE structural zero |
| D — Autres | 668x | **Grand Livre (GL xls), tous journaux** | Info only — excluded from TOTAL |

`TOTAL COMPTABILITE = Section A + B + C` (Section D excluded)

---

## RULE 5 — Mandatory Excel Formulas in Feuil2

**Never write Python-computed values in formula cells.**

| Cell | Formula |
|------|---------|
| Col V (CF/P+FNE), row r | `=T{r}+U{r}` |
| Col Y (TOTAL), row r | `=R{r}+S{r}+V{r}+W{r}+X{r}` ← uses V, not T+U separately |
| TOTAL PAIE, each col | `=SUM(col{start}:col{end})` |
| Subtotals A, B, C | `=SUM(col{section_start}:col{section_end})` |
| TOTAL COMPTA | `={col}{stA}+{col}{stB}+{col}{stC}` |
| ECART | `={col}{total_compta}-{col}{total_paie}` |

Values in individual account rows (661110, 664120, etc.) may remain numeric.

---

## RULE 6 — Auto-trigger Gap Analysis After Feuil2 Rebuild

**After any rebuild of Feuil2, automatically run the full summarise workflow.**

Never do partial Feuil1 updates. Always:
1. Save `row_total_paie`, `row_total_compta`, `row_ecart` to `.audit-session.json`
2. Call `build_feuil1_summary.py --write` (passing row numbers)
3. In interactive mode: show ecart table and ask user to validate before writing Feuil1
4. In unattended mode: run automatically and report completion

`build_feuil1_summary.py` always reads values directly from Feuil2 — never from session cache.
Feuil1 is always rewritten from scratch — never partially updated.

---

## Regression Reference — CIFM Cameroun, Exercice 2025

After a correct full run, verify these totals:

| Metric | Expected value |
|--------|---------------|
| Nb employés (PAIE) | 174 |
| TOTAL PAIE — Col R (SAL BRUT) | ~651 500 000 FCFA |
| TOTAL PAIE — Col Y (TOTAL général) | **666 286 638 FCFA** |
| TOTAL COMPTA — Col Y | **682 630 194 FCFA** |
| ECART — Col U (FNE) | **−11 953 704 FCFA** (structural) |
| ECART — Col Y (TOTAL) | **+16 343 556 FCFA** |
| Nb comptes Section A (661x+663x) | 13 |
| Source Section A | Balance Générale (BG xlsx) |
| Source Sections B/C/D | Grand Livre (GL xls), tous journaux |

---

## CSV Parsing Notes

Both CSV files (`LIVREPAIE*.CSV`, `CHARGESPATRONALES*.CSV`) use Excel text-forcing format:
- Delimiter: `;`
- Encoding: `latin-1`
- Values wrapped in `="value"` or `=""value""`
- French decimal: `,` inside quotes (e.g. `87453","00` → `87453.00`)

Never use `pandas.read_csv()` directly — use the line-by-line parser in `scripts/parse_*.py`.

Matricule validation regex: `re.match(r'^\d{3,}', matricule)` — must start with ≥3 digits.

BRUT code: exact match `TypeCode == "BRUT"` only.
