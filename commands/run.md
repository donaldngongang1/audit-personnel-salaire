---
name: run
description: "Run the full payroll audit workflow — detect files, extract data, build reconciliation, verify, and summarise. Bilingual (FR/EN). Starts with a mode choice (interactive or unattended)."
argument-hint: "[--lang fr|en] [--dir /path/to/audit/folder]"
allowed-tools: Bash, Read, Write, Edit, AskUserQuestion, TodoWrite
---

## /audit-personnel-salaire:run

Execute the complete payroll/personnel charges audit workflow from source files to final workpaper.

### Step 1 — Language & Mode Selection

If `.audit-session.json` does not exist in the working directory, use AskUserQuestion with:

**Q1 — Language / Langue:**
- Option A: "Français / French"
- Option B: "English / Anglais"

**Q2 — Run mode / Mode d'exécution:**
- Option A: "Interactif / Interactive — pause à chaque étape pour confirmation (Recommandé)"
  - Description: "Claude pauses after each step and asks for confirmation before continuing"
- Option B: "Non supervisé / Unattended — poser les questions au départ, puis exécuter automatiquement"
  - Description: "Ask all upfront questions, then run all steps automatically to completion"

Store answers in `.audit-session.json`.

### Step 2 — Source File Detection

Run `scripts/detect_files.py --dir <working_dir>` and present results as a table.

For each missing or ambiguous file, use AskUserQuestion to confirm path.
If no `FT-P-2*.xlsx` found, copy `assets/FT-P-2-template.xlsx` to the working directory.
Ask user: "Quel nom donner au fichier de travail? / What name for the working file?"
Default: `FT-P-2- Exhaustivité des charges du personnel (Salaire brut).xlsx`

### Step 3 — Package Check

Run `scripts/check_packages.py`. If any package is missing, show exact `pip install` command
and pause (even in unattended mode) until user confirms installation.

### Step 4 — Data Extraction

Run in sequence (or parallel if unattended):
```bash
python scripts/parse_balance.py
python scripts/parse_grand_livre.py
python scripts/parse_charges_patronales.py
python scripts/parse_livre_paie.py
python scripts/build_extracts.py
```

In interactive mode: after each extraction, show row count and ask "Continue? / Continuer?"
In unattended mode: run all at once, report summary at end.

### Step 5 — Reconciliation Build

```bash
python scripts/build_reconciliation.py --section all
```

In interactive mode: after PAIE section, pause and show TOTAL PAIE figures for review.
After COMPTA section, pause and show TOTAL COMPTA figures.

### Step 6 — Verification

```bash
python scripts/eval_totals.py
python scripts/eval_ecart.py
python scripts/eval_formulas.py
```

Show results. If any check fails, pause and ask user whether to re-run fix or investigate manually.

### Step 7 — Gap Analysis & Feuil1

```bash
python scripts/build_feuil1_summary.py
```

Before running, ask user for any metadata needed for Feuil1 header (Société, Auditeur, Période)
if not already in `.audit-session.json`.

Present each ecart with explanation. In interactive mode, allow user to edit explanations.

### Step 8 — Completion

Print final summary:
```
✅ Audit complet / Audit complete
   Fichier: FT-P-2-....xlsx
   Employés PAIE: N
   Comptes COMPTA: M
   Écarts: [NONE / X écarts — voir Feuil1]
```

Update `.audit-session.json` with `"steps_completed": ["all"]`.

---

**Note**: All scripts read file paths from `.audit-session.json`. Always run `detect-files` first
if the session file does not exist.
