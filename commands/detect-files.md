---
name: detect-files
description: "Scan the current directory for audit source files (Balance Générale, Grand Livre, Livre de Paie, Charges Patronales, FT-P-2). Confirm paths with user and save to .audit-session.json."
argument-hint: "[--dir /path/to/scan]"
allowed-tools: Bash, AskUserQuestion, Write
---

## /audit-personnel-salaire:detect-files

Scan for all required source files and confirm them with the user.

### Execution

```bash
python scripts/detect_files.py --dir "${ARGS[--dir]:-$(pwd)}"
```

### After script output

Present the results table:
```
📁 Scan Results — [directory]
┌──────────────────────────┬────────┬──────────────────────────────────┐
│ File Type                │ Status │ Path                             │
├──────────────────────────┼────────┼──────────────────────────────────┤
│ Balance Générale         │   ✅   │ BG CIFM 2025.xlsx                │
│ Grand Livre Général      │   ✅   │ GRAND LIVRE GENERALE CIFM 2025.. │
│ Livre de Paie            │   ✅   │ LIVREPAIECIFM2025.CSV            │
│ Charges Patronales       │   ✅   │ CHARGESPATRONALESCIFM2025.CSV    │
│ Feuille de Travail (FT)  │   ✅   │ FT-P-2- Exhaustivité....xlsx     │
└──────────────────────────┴────────┴──────────────────────────────────┘
```

For each ❌ Not Found file, use AskUserQuestion:
"[File type] not found. Please provide the path, or type 'skip' if not applicable."

For each ⚠️ Multiple matches, use AskUserQuestion with each match as an option.

After confirmation, save `.audit-session.json` with confirmed paths.

If FT-P-2 not found: copy `assets/FT-P-2-template.xlsx` to the working directory.
Ask: "Comment nommer le fichier de travail? / What name for the working file?"
