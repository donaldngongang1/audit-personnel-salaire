---
name: Gap Analysis
description: >
  Activates when the user asks to "explain the gap", "explain ecart", "why is there a difference",
  "analyse l'écart", "pourquoi y a-t-il un écart", "investigate discrepancies", "what caused the
  difference", "fill Feuil1 summary", "remplir Feuil1", "summarise audit findings", or after
  reconciliation completes with non-zero ecarts. Reads ALL values directly from Feuil2 in the
  workbook (not from session cache), generates bilingual explanations, and rewrites Feuil1 ENTIRELY.
version: 1.1.0
---

## Gap Analysis — Skill Guide (v1.1 — read from Feuil2, full Feuil1 rewrite)

After Feuil2 is built and ecarts are computed, this skill investigates each ecart and rewrites
Feuil1 completely from scratch. Never do partial Feuil1 updates.

---

### Step 1 — Read Values from Feuil2 (Not from Session Cache)

**Always read directly from the Feuil2 sheet** in the workbook. Do not rely on `.audit-session.json`
for ecart amounts — the file may be stale.

1. Open the workbook (read_only=True, data_only=True to get computed values, not formulas).
2. Scan column A (or B) for labels to locate:
   - `row_total_paie`: row where label contains "TOTAL PAIE"
   - `row_total_compta`: row where label contains "TOTAL COMPTABILITE"
   - `row_ecart`: row where label contains "ECART" and "TOTAL"
3. Read values for columns R(18) through Y(25) from each of those rows.
4. Compute `ecart_per_col[c] = compta_val[c] − paie_val[c]` for each column.

---

### Step 2 — Ecart Classification

For each column with non-zero ecart:

| Column | Expected ecart | Classification |
|--------|---------------|---------------|
| U (FNE) | = −PAIE_FNE total | Structurel: FNE non comptabilisé en GL |
| R (SAL BRUT) | varies | Différence BG 661x+663x vs livre de paie |
| S (CNPS/P) | near 0 | Vérifier 664120 vs charges patronales CNPS |
| T (CF/P) | near 0 | Vérifier 664380 vs charges patronales CF/P |
| W (AF) | near 0 | Vérifier 664110 vs charges patronales AF |
| X (AT) | near 0 | Vérifier 664130 vs charges patronales AT |

Rounding: |ecart| ≤ 1 FCFA → "Arrondi" classification, no further investigation needed.
Significant: |ecart| > 1 000 FCFA → flag for investigation.

---

### Step 3 — Generate Bilingual Explanations

For each column ecart, generate both FR and EN explanations:

**FNE (column U — structural zero in GL):**
> FR: "Le FNE (Fond National de l'Emploi) n'est pas comptabilisé dans le Grand Livre Général (compte 66x). C'est une retenue sur salaire versée directement à l'ONEM sans écriture comptable distincte dans la classe 66. Écart structurel normal — aucune action requise."
> EN: "FNE contributions are salary deductions paid directly to ONEM. They are not recorded as a Class 66 entry in the General Ledger. This is a normal structural gap — no corrective action needed."

**Section A (661x+663x) gap:**
> FR: "Écart entre le salaire brut du livre de paie ([PAIE] FCFA) et les charges comptabilisées en 661/663 ([COMPTA] FCFA). Différence = [ECART] FCFA. À investiguer : vérifier si des primes ou régularisations ont été comptabilisées hors période de paie."

---

### Step 4 — Interactive Mode: Review Before Writing

In interactive mode, before writing Feuil1:

1. Display the ecart table:
```
| Colonne | TOTAL PAIE   | TOTAL COMPTA | ECART         | Classification        |
|---------|-------------|-------------|---------------|----------------------|
| R BRUT  | 651,234,567 | 665,578,123 | +14,343,556   | À investiguer        |
| S CNPS  |  56,789,012 |  56,789,012 |             0 | ✅ Équilibré          |
| T CF/P  |  12,345,678 |  12,345,678 |             0 | ✅ Équilibré          |
| U FNE   |  11,953,704 |           0 | −11,953,704   | Structurel (FNE)     |
| V CF+FN |  24,299,382 |  12,345,678 | −11,953,704   | Structurel            |
| W AF    |   8,123,456 |   8,123,456 |             0 | ✅ Équilibré          |
| X AT    |   2,654,321 |   2,654,321 |             0 | ✅ Équilibré          |
```

2. Use AskUserQuestion: "Ces écarts sont-ils corrects? Voulez-vous modifier les justifications avant l'écriture de Feuil1? / Are these gaps correct? Do you want to edit the explanations before writing Feuil1?"

3. If user wants to edit: allow custom text input for each explanation.

---

### Step 5 — Rewrite Feuil1 Completely

**NEVER do partial updates to Feuil1.** Always rewrite the entire sheet from scratch to ensure consistency.

When called after reconciliation-builder, receive these row coordinates from session:
- `row_total_paie`, `row_total_compta`, `row_ecart`

Steps:
1. Delete all content from Feuil1 (keep sheet, remove all cell values and styles).
2. Rebuild the complete Feuil1 structure:
   - Title / header block
   - Metadata (Société, Exercice, Période, Auditeur)
   - Summary table: PAIE totals, COMPTA totals, ECART per column
   - Gap explanations for each non-zero ecart
   - Overall status: ✅ Réconcilié / ⚠️ Écarts justifiés / ❌ Écarts inexpliqués
3. Save the workbook.

---

### Step 6 — Update Session State

After writing Feuil1, update `.audit-session.json`:
```json
{
  "gap_analysis": {
    "total_paie_R": 651234567,
    "total_compta_R": 665578123,
    "ecart_R": 14343556,
    "ecart_U": -11953704,
    "explanations": { "U": "FNE structurel...", "R": "..." },
    "overall_status": "ecarts_justifies"
  },
  "feuil1_written": true,
  "steps_completed": ["all"]
}
```

---

### Regression Reference (CIFM 2025)

Expected ecart values after correct reconciliation:
- ECART R (SAL BRUT) = +16 343 556 FCFA (COMPTA > PAIE — provisions/régularisations)
- ECART U (FNE) = −11 953 704 FCFA (structural — FNE hors GL)
- ECART Y (TOTAL) = +16 343 556 FCFA
- All other columns: 0 FCFA
