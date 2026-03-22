---
name: Reconciliation Builder
description: >
  Activates when the user asks to "build Feuil2", "fill the reconciliation sheet", "remplir le
  rapprochement", "créer le tableau de bord", "reconcile payroll and accounting", "fill paie
  section", "fill compta section", "build ecart row", "construire Feuil2", or when the
  reconciliation step of the audit workflow is reached. Provides the complete logic for
  populating every row and column of the Feuil2 workpaper, then automatically triggers the
  full gap-analysis (summarise) workflow.
version: 1.1.0
---

## Reconciliation Builder — Skill Guide (v1.1 — SYSCOHADA + Excel formulas)

Feuil2 is the central reconciliation workpaper. After building it, this skill MUST automatically
trigger the full `summarise` (gap-analysis) skill/command.

---

### Column Map (SYSCOHADA)

| Col | Letter | Header | PAIE source | COMPTA source |
|-----|--------|--------|-------------|---------------|
| 14 | N | Salaire Base | **ALWAYS BLANK** | **ALWAYS BLANK** |
| 15 | O | Ancienneté | **ALWAYS BLANK** | **ALWAYS BLANK** |
| 16 | P | H. Sup | **ALWAYS BLANK** | **ALWAYS BLANK** |
| 17 | Q | Autre Gain | **ALWAYS BLANK** | **ALWAYS BLANK** |
| 18 | R | SAL BRUT | LivrePaie BRUT | BG — 661x+663x net solde |
| 19 | S | CNPS/P | ChargesPatronales 4500 | GL — compte 664120 |
| 20 | T | CF/P | ChargesPatronales 4100 | GL — compte 664380 |
| 21 | U | FNE | ChargesPatronales 4400 | 0 (non comptabilisé GL) |
| 22 | V | CF/P+FNE | formula | formula |
| 23 | W | AF | ChargesPatronales 4800 | GL — compte 664110 |
| 24 | X | AT | ChargesPatronales 4900 | GL — compte 664130 |
| 25 | Y | TOTAL | formula | formula |

---

### Source Rules by Section (CRITICAL)

**Section A — Rémunérations directes (661x + 663x):**
→ Source = **Balance Générale (BG xlsx)**, col4=MvtDebit, col5=MvtCredit
→ Solde Net = col4 − col5
→ Written to column **R**
→ Reason: BG is exhaustive; GL may be missing entries for some accounts

**Section B — Cotisations CNPS (664110, 664120, 664130):**
→ Source = **Grand Livre (GL xls)**, ALL journals (no journal filter)
→ 664120 → col S (CNPS/P), 664110 → col W (AF), 664130 → col X (AT)

**Section C — CF/P et FNE:**
→ Source = **Grand Livre (GL xls)**, ALL journals
→ 664380 → col T (CF/P)
→ FNE → col U = **0** (structural: FNE is a salary deduction, not in GL 66x)

**Section D — Autres charges sociales (668x) — INFORMATIVE ONLY:**
→ Source = **Grand Livre (GL xls)**, ALL journals
→ Written for information only
→ **NOT included in TOTAL COMPTABILITE**

---

### Mandatory Excel Formulas (NOT hardcoded values)

Every formula cell must be an actual Excel formula string — not a Python-computed number:

**Employee rows (row r, PAIE section):**
```
col V : =T{r}+U{r}
col Y : =R{r}+S{r}+V{r}+W{r}+X{r}
```
Note: Y uses V (not T+U separately) — so V must be written before Y on the same row.
Cols R, S, T, U, W, X: write numeric values from payroll pivot CSVs (acceptable).

**TOTAL PAIE row:**
```
col R : =SUM(R{paie_start}:R{paie_end})
col S : =SUM(S{paie_start}:S{paie_end})
... (same for T, U, V, W, X)
col Y : =SUM(Y{paie_start}:Y{paie_end})
```

**Subtotal A row (after Group A accounts):**
```
col R : =SUM(R{groupA_start}:R{groupA_end})
col Y : =SUM(Y{groupA_start}:Y{groupA_end})   ← if applicable
```

**TOTAL COMPTABILITE row:**
```
col R : =R{subtot_A}+R{subtot_B}+R{subtot_C}
col S : =S{subtot_A}+S{subtot_B}+S{subtot_C}
... (same for T, U, V, W, X, Y)
```
(Section D excluded from this formula.)

**ECART row:**
```
col R : =R{total_compta}-R{total_paie}
col S : =S{total_compta}-S{total_paie}
... (same for T, U, V, W, X, Y)
```

---

### Column Blanking Rule (UNCHANGED)

Columns N, O, P, Q (14–17) MUST be None/blank on EVERY row — PAIE and COMPTA sections.
Never write values or formulas to these columns.

---

### Session State After Build

After completing Feuil2, save to `.audit-session.json`:
```json
{
  "feuil2_build": {
    "row_total_paie": 178,
    "row_total_compta": 214,
    "row_ecart": 217,
    "total_paie": 666286638,
    "total_compta": 682630194,
    "ecart": 16343556,
    "source_section_A": "BG",
    "source_sections_BCD": "GL"
  }
}
```

---

### Auto-trigger Gap Analysis (MANDATORY)

After Feuil2 is complete and saved:

1. Store `row_total_paie`, `row_total_compta`, `row_ecart` in session.
2. **Automatically call the `summarise` skill** (full gap-analysis):
   - In **interactive mode**: display ECART values per column, ask user to validate before writing Feuil1.
   - In **unattended mode**: run summarise automatically, report completion in final summary.
3. The reconciliation-builder skill MUST call `summarise` at the end — partial Feuil1 updates are not permitted.

---

### Unmerge Before Writing

Before writing to any row ≥ COMPTA_START, unmerge all merged cells in that region:
```python
to_unmerge = [mr for mr in list(ws.merged_cells.ranges) if mr.min_row >= start_row]
for mr in to_unmerge:
    ws.unmerge_cells(str(mr))
```

---

### Regression Reference (CIFM 2025)

After a successful build, verify these expected totals:
- TOTAL PAIE   = 666 286 638 FCFA (174 employees)
- TOTAL COMPTA = 682 630 194 FCFA (13 accounts 661x+663x + CNPS + CF/P)
- ECART        =  16 343 556 FCFA
