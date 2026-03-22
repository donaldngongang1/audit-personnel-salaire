---
name: Audit Verification
description: >
  Activates when the user asks to "verify calculations", "run evals", "check my work", "vérifier
  les calculs", "recalculer les totaux", "contrôler le rapprochement", "run audit checks",
  "validate figures", "check for errors in Feuil2", "audit the audit", or when the post-write
  verification hook fires after any Excel file is modified. Independently recalculates all totals
  and ecarts from source data and compares them to what was written in the workbook.
version: 1.0.0
---

## Audit Verification — Skill Guide

This skill runs three independent verification scripts after each write to the workbook.
Financial data is sensitive — never skip verification.

---

### Three Verification Scripts

| Script | What it verifies | Pass condition |
|--------|-----------------|---------------|
| `eval_totals.py` | PAIE and COMPTA totals in Feuil2 match re-computed values from source CSVs/Excel | All column totals within 1 FCFA tolerance |
| `eval_ecart.py` | ECART row = TOTAL COMPTA − TOTAL PAIE for each column | Every ecart cell equals the difference |
| `eval_formulas.py` | Y column formula = R+S+T+U+W+X; V = T+U; no stale `=N+O+P+Q` references remain | All formula strings correct; N/O/P/Q cells = None |

---

### eval_totals.py Logic

1. Re-read source files (using same parsers as extraction).
2. Re-compute expected column totals independently:
   - PAIE_R = sum of all employee SAL BRUT from LivrePaie CSV
   - PAIE_S = sum of all employee CNPS from ChargesPatronales CSV
   - PAIE_T = sum of all CF/P, PAIE_U = sum of all FNE, PAIE_W = sum of AF, PAIE_X = sum of AT
   - COMPTA_R = sum of MvtDebit − MvtCredit for accounts 661–663 from Balance Générale
   - COMPTA_S = net solde of account 664120 (CNPS AV)
   - COMPTA_T = net solde of account 664380 (Prov. CF/P)
   - COMPTA_U = 0 (structural)
   - COMPTA_W = net solde of account 664110 (CNPS AF)
   - COMPTA_X = net solde of account 664130 (CNPS AT)
3. Compare to workbook TOTAL PAIE and TOTAL COMPTA cells.
4. Report: ✅ OK / ❌ MISMATCH with expected vs actual amounts.

---

### eval_ecart.py Logic

1. Read TOTAL PAIE row cell values from workbook (by row number stored in session state).
2. Read TOTAL COMPTA row cell values.
3. Read ECART row cell values.
4. For each column: assert `ecart == compta − paie` within 1 FCFA tolerance.
5. Report columns that pass/fail.

---

### eval_formulas.py Logic

1. Scan every cell in columns N, O, P, Q (14–17) on all data rows.
2. Assert each cell value is None or 0.
3. Scan every Y cell (col 25) on PAIE rows: assert formula contains `R+S+T+U+W+X`.
4. Scan every V cell (col 22): assert formula contains `T+U`.
5. Scan for any cell containing `=N` or `=O` or `=P` or `=Q` in its formula.
6. Report: list of cells with violations.

---

### PostToolUse Hook Integration

After every Write or Edit tool use that touches a `.xlsx` file, the `post-write-verify` hook fires.
The hook runs all three eval scripts and outputs a summary like:
```
🔍 Audit Verification (post-write)
✅ eval_totals: All 10 column totals match source data
✅ eval_ecart: All ecart cells correct
✅ eval_formulas: No stale N/O/P/Q references found
```
If any script fails, the hook outputs a ❌ warning with the specific mismatches and asks the user
whether to re-run the fix scripts.

---

### Tolerance Policy

- Rounding tolerance: **1 FCFA** (amounts are integers; CSV parsing may produce float artefacts).
- If the difference is between 1 and 1,000 FCFA: flag as "Écart d'arrondi probable".
- If the difference exceeds 1,000 FCFA: flag as "Écart significatif — investigation requise".

---

### Continuous Verification During Build

During `build_reconciliation.py`, call eval scripts after each section:
1. After writing PAIE rows → run `eval_totals.py --section paie`
2. After writing COMPTA rows → run `eval_totals.py --section compta`
3. After writing ECART row → run `eval_ecart.py` and `eval_formulas.py`

This ensures errors are caught at the earliest possible stage.
