---
name: verify
description: "Run all three audit verification scripts (eval_totals, eval_ecart, eval_formulas) and display a pass/fail report. Safe to run at any time on the current workbook."
argument-hint: "[--script totals|ecart|formulas|all]"
allowed-tools: Bash, Read
---

## /audit-personnel-salaire:verify

Independently verify all figures in the FT-P-2 workbook against source data.

### Pre-check

Read `.audit-session.json` for workbook path and source file paths.

### Run Verification Scripts

```bash
python scripts/eval_totals.py
python scripts/eval_ecart.py
python scripts/eval_formulas.py
```

Or individually based on `--script` argument.

### Output Format

```
🔍 Audit Verification Report
════════════════════════════════════════════════════════
eval_totals   ✅ PASS — All 10 column totals match source data
eval_ecart    ✅ PASS — All ecart cells = COMPTA − PAIE
eval_formulas ✅ PASS — No stale N/O/P/Q references; Y=R+S+T+U+W+X

Overall: ✅ ALL CHECKS PASSED
════════════════════════════════════════════════════════
```

If any check fails:
```
eval_totals   ❌ FAIL
  Row 178 (TOTAL PAIE), Col R: Expected 2,456,782,103 | Found 2,456,782,000 | Diff: 103
```

### On Failure

Present the specific mismatches. Offer to re-run the reconciliation builder to fix them.
Financial data is sensitive — never ignore verification failures without explicit user approval.
