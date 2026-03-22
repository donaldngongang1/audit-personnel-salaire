---
name: Gap Analysis
description: >
  Activates when the user asks to "explain the gap", "explain ecart", "why is there a difference",
  "analyse l'écart", "pourquoi y a-t-il un écart", "investigate discrepancies", "what caused the
  difference", "fill Feuil1 summary", "remplir Feuil1", "summarise audit findings", or after
  reconciliation completes with non-zero ecarts. Systematically investigates each ecart column,
  identifies root causes, and prepares explanations for the Feuil1 summary sheet.
version: 1.0.0
---

## Gap Analysis — Skill Guide

After Feuil2 is built and ecarts are computed, this skill drives the investigation and explanation
of each non-zero ecart. Results feed into the Feuil1 summary sheet.

---

### Known Structural Gaps (Always Explain)

Some ecarts are expected and always need the same explanation:

| Column | Typical gap cause | Standard explanation (FR) | Standard explanation (EN) |
|--------|-------------------|--------------------------|--------------------------|
| U (FNE) | FNE is paid but not booked in GL | Le FNE ne fait pas l'objet d'une écriture comptable distincte dans le GL | FNE contributions are paid externally and not recorded in the general ledger |
| Any 66x | Timing difference: payroll month ≠ GL period | Écart de période : les salaires du mois de [M] sont comptabilisés en [M+1] | Timing difference: [month] payroll booked in following period |
| R | Gross pay difference | Différence entre le brut comptabilisé (661xxx) et le brut de paie | Difference between accounting gross (661xxx) and payroll gross |

---

### Gap Investigation Algorithm

For each non-zero ecart column:

1. **Check magnitude**: Small rounding (<= 100 FCFA) → mark as "Écart d'arrondi / Rounding difference"

2. **Check FNE column (U)**: Always zero on COMPTA side → ecart = −PAIE_FNE total
   → Explanation: "FNE non comptabilisé au Grand Livre"

3. **Compare account totals**: Sum all GL accounts for the column's mapping vs PAIE total.
   Identify which account(s) contribute to the gap.

4. **Check for missing accounts**: Are there accounts in PAIE charges not matched in GL?
   (e.g., employee categories present in payroll but missing from chart of accounts)

5. **Check for period mismatch**: Compare payroll month in CSV filename with GL transaction dates.

6. **Check for double entries**: Look for accounts booked in both Group A and Group B (rare).

---

### Feuil1 Summary Sheet — Structure

Feuil1 is the executive summary page. Never change its formatting; only fill in the value cells.

Typical Feuil1 sections to populate:
- **Header**: Société, Exercice, Période, Auditeur (ask user if not in session state)
- **Synthèse des charges**: Totals from PAIE (SAL BRUT, CNPS, CF/P, FNE, AF, AT, TOTAL)
- **Synthèse comptable**: Totals from COMPTABILITE section
- **Écarts**: Per-column ecart amounts and explanations
- **Observations**: Free-text audit observations from gap analysis
- **Conclusion**: Overall reconciliation status (✅ Aucun écart / ⚠️ Écarts justifiés / ❌ Écarts inexpliqués)

---

### Explanation Quality Standards

Each gap explanation must include:
1. **Amount**: Exact FCFA amount of the ecart
2. **Direction**: PAIE > COMPTA or COMPTA > PAIE
3. **Root cause**: One of: timing, missing GL entry, rounding, scope difference, data error
4. **Supporting evidence**: Which accounts / employees drive the gap
5. **Audit recommendation**: Book missing entry / Adjust payroll / No action needed

---

### Script: build_feuil1_summary.py

Reads `.audit-session.json` for the workbook path.
Reads ecart values from Feuil2 ECART row (dynamically located by scanning for "ECART" label).
Fills Feuil1 values using openpyxl named ranges or fixed cell coordinates stored in session state.
Does NOT change any formatting or formula on Feuil1 — values only.

---

### Bilingual Output

All gap explanations are stored in both languages in `.audit-session.json`:
```json
{
  "ecarts": {
    "U": {
      "amount": -11953704,
      "explanation_fr": "Le FNE (Fond National de l'Emploi) ne fait pas l'objet d'une comptabilisation distincte dans le Grand Livre Général. Les charges FNE sont calculées sur la base du livre de paie mais versées directement sans écriture.",
      "explanation_en": "FNE (National Employment Fund) contributions are not recorded as a separate entry in the General Ledger. They are calculated from the payroll register and paid directly without a GL booking.",
      "category": "hors_champ_gl",
      "action": "information_only"
    }
  }
}
```
