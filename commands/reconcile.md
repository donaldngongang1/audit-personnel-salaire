---
name: reconcile
description: "Build Feuil2 reconciliation workpaper: fill PAIE section (employee rows), COMPTABILITE section (GL accounts), TOTAL rows, and ECART row."
argument-hint: "[--section paie|compta|all]"
allowed-tools: Bash, Read, AskUserQuestion
---

## /audit-personnel-salaire:reconcile

Build the Feuil2 reconciliation workpaper from the Extract sheets.

### Pre-check

Read `.audit-session.json`. Verify `steps_completed` includes `extract`.
If not: "Please run /extract first. / Veuillez d'abord exécuter /extract."

Close Excel before writing (Windows):
```bash
tasklist | grep -i EXCEL && echo "⚠️ Excel is open — please close it before continuing."
```
Use AskUserQuestion: "Excel appears to be open. Please close FT-P-2 and press 'Continue'."

### Build Reconciliation

```bash
python scripts/build_reconciliation.py --section "${ARGS[--section]:-all}"
```

### Interactive Checkpoints (when run_mode = interactive)

1. After PAIE section: show TOTAL PAIE per column (R→Y). Ask: "Review the PAIE totals above. Continue to COMPTABILITE? / Vérifiez les totaux PAIE. Continuer vers COMPTABILITE?"

2. After COMPTA section: show TOTAL COMPTA. Ask: "Review the COMPTA totals above. Build ECART row? / Vérifiez les totaux COMPTA. Construire la ligne ECART?"

3. After ECART row: run verification immediately.

### Post-write Verification

Always run after build completes:
```bash
python scripts/eval_totals.py
python scripts/eval_ecart.py
python scripts/eval_formulas.py
```

Show verification result. If any check fails, present the mismatch details and ask:
"A verification error was found. Run auto-fix? / Une erreur de vérification a été trouvée. Lancer la correction automatique?"

Update `steps_completed` in `.audit-session.json`.
