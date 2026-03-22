---
name: Reconciliation Builder
description: >
  Activates when the user asks to "build Feuil2", "fill the reconciliation sheet", "remplir le
  rapprochement", "créer le tableau de bord", "reconcile payroll and accounting", "fill paie section",
  "fill compta section", "build ecart row", "construire Feuil2", or when the reconciliation step
  of the audit workflow is reached. Provides the complete logic for populating every row and
  column of the Feuil2 workpaper from the Extract sheets.
version: 1.0.0
---

## Reconciliation Builder — Skill Guide

Feuil2 is the central reconciliation workpaper. It has two main sections — PAIE and COMPTABILITE —
plus a TOTAL and ECART row. The script `scripts/build_reconciliation.py` implements all of this.

---

### Feuil2 Column Map

| Col | Letter | Header | Source |
|-----|--------|--------|--------|
| 1–13 | A–M | Employee/account identity fields | Copied from payroll data |
| 14 | N | Salaire Base | **ALWAYS BLANK** |
| 15 | O | Ancienneté | **ALWAYS BLANK** |
| 16 | P | H. Sup | **ALWAYS BLANK** |
| 17 | Q | Autre Gain | **ALWAYS BLANK** |
| 18 | R | SAL BRUT | LivrePaie BRUT total per employee; for COMPTA = net solde of accounts 661-663 |
| 19 | S | CNPS/P (Pension AV) | ChargesPatronales code 4500; for COMPTA = GL account 664120 |
| 20 | T | CF/P | ChargesPatronales code 4100; for COMPTA = GL account 664380 |
| 21 | U | FNE | ChargesPatronales code 4400; for COMPTA = 0 (FNE not in GL) |
| 22 | V | CF/P + FNE | Formula: `=T{r}+U{r}` |
| 23 | W | AF | ChargesPatronales code 4800; for COMPTA = GL account 664110 |
| 24 | X | AT | ChargesPatronales code 4900; for COMPTA = GL account 664130 |
| 25 | Y | TOTAL | Formula: `=R{r}+S{r}+T{r}+U{r}+W{r}+X{r}` |

**CRITICAL**: Columns N, O, P, Q (14–17) MUST be blank (value=None) on EVERY data row.
Never write formulas or values to these columns.

---

### Section I — PAIE (Employee Rows)

- One row per employee from the merged LivrePaie + ChargesPatronales pivot.
- Sort by Matricule ascending.
- Row start: dynamically determined (look for "TOTAL PAIE" label to find the end).
- Write R (SAL BRUT), S (CNPS/P), T (CF/P), U (FNE), W (AF), X (AT) from pivot data.
- V = formula `=T{r}+U{r}`, Y = formula `=R{r}+S{r}+T{r}+U{r}+W{r}+X{r}`
- Alternating row fills: even rows = `F5F8FC`, odd rows = white/None.

**TOTAL PAIE row**: `=SUM(col{DATA_START}:col{TOT_PAIE_ROW-1})` for each of R→X, Y.
Style: `D6E4F0` fill, bold `1F4E79` font, medium top+bottom border.

---

### Section II — COMPTABILITE (GL Account Rows)

Four groups of accounts, each with a subtotal row:

**Group A — Salaires et appointements (accounts 661–663)**
Write net solde (MvtDebit − MvtCredit from Balance Générale, filtered by account) into **column R**.
Accounts: 661110, 661120, 661130, 661200, 661210, 661220, 661300, 661380, 661410, 661800, 663101, 663102, 663410.

**Group B — CNPS (accounts 664110, 664120, 664130)**
- 664120 (CNPS Pension AV) → **column S**
- 664110 (CNPS AF) → **column W**
- 664130 (CNPS AT) → **column X**

**Group C — CF/P et FNE (accounts 664380, FNE_NA)**
- 664380 (Prov. CF/P) → **column T**
- FNE_NA (FNE non comptabilisé) → **column U = 0** ← audit finding!

**Group D — Informations (accounts 668420, 668430, 668700)**
- Written for information only; value = 0 or GL amount.
- **Excluded from TOTAL COMPTABILITE** — shown in grey.

**TOTAL COMPTABILITE**: `=GroupA_subtotal + GroupB_subtotal + GroupC_subtotal` (D excluded).

---

### ECART Row (Section III)

For each column R through X and Y:
```excel
=TOTAL_COMPTA_cell − TOTAL_PAIE_cell
```

Style: `843C0C` dark orange fill, white bold font, medium borders.

An ecart of 0 = perfect reconciliation. Non-zero ecart = audit finding to be explained.

---

### Styling Standards

- Data rows: Arial 10pt, thin borders (`D9D9D9`), right-aligned numeric cells.
- Total rows: Arial 10pt bold, `1F4E79` blue, medium borders (`1F4E79`), `D6E4F0` or `C6EFCE` fill.
- ECART row: `843C0C` fill, white bold font.
- Number format for all numeric cells: `#,##0;(#,##0);"-"`

---

### Unmerge Before Writing

Before writing to Feuil2, always unmerge any merged cells in the region being written:
```python
to_unmerge = [mr for mr in list(ws.merged_cells.ranges) if mr.min_row >= start_row]
for mr in to_unmerge:
    ws.unmerge_cells(str(mr))
```

---

### build_reconciliation.py Entry Point

The script reads `.audit-session.json` for file paths and `steps_completed` to resume if needed.
It accepts `--section paie|compta|all` to run partial rebuilds.
After writing, it calls `eval_totals.py` and `eval_ecart.py` for immediate verification.
