---
name: summarise
description: "Analyse ecarts, generate gap explanations, and fill Feuil1 summary sheet. Asks for Société, Auditeur, and Période if not already stored. Bilingual (FR/EN)."
argument-hint: "[--lang fr|en]"
allowed-tools: Bash, Read, Write, AskUserQuestion
---

## /audit-personnel-salaire:summarise

Analyse gaps and fill the Feuil1 executive summary sheet.

### Pre-check

Read `.audit-session.json`. Verify `steps_completed` includes `reconcile` and `verify`.
If verification has not passed: warn and ask user to confirm they want to proceed anyway.

### Collect Feuil1 Metadata

Use AskUserQuestion for any fields missing from `.audit-session.json`:
- "Nom de la société / Company name?"
- "Nom de l'auditeur / Auditor name?"
- "Période d'audit (ex: Janvier-Décembre 2025) / Audit period?"
- "Date du rapport / Report date?" (default: today)

Store in `.audit-session.json`.

### Run Gap Analysis

```bash
python scripts/build_feuil1_summary.py
```

Script reads Feuil2 ECART row, matches to known gap patterns, and generates explanations.

### Review Explanations (interactive mode)

For each non-zero ecart column, show:
```
Column U (FNE): Écart = −11,953,704 FCFA
  Explication FR: Le FNE ne fait pas l'objet d'une comptabilisation...
  Explanation EN: FNE contributions are not recorded...
  Category: hors_champ_gl | Action: information_only
```

Ask: "Souhaitez-vous modifier cette explication? / Do you want to edit this explanation?"

### Final Feuil1 Fill

Run `build_feuil1_summary.py --write` to write values to Feuil1.

### Completion

```
✅ Feuil1 complétée / Feuil1 completed
   Société: [name]
   Période: [period]
   Auditeur: [name]
   Écarts: X colonne(s) avec écart / X column(s) with gap
   Statut: ✅ Réconcilié / ⚠️ Écarts justifiés / ❌ Écarts non expliqués
```

Ask: "Souhaitez-vous ouvrir le fichier? / Do you want to open the file?"
If yes: `python -c "import subprocess,json; d=json.load(open('.audit-session.json')); subprocess.Popen(['start',d['files']['feuille_travail']], shell=True)"`
