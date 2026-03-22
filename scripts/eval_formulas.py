"""
eval_formulas.py — Verify formula integrity in Feuil2:
  1. Columns N/O/P/Q (14-17) must be blank on all data rows.
  2. Y column formula must be =R+S+T+U+W+X.
  3. V column formula must be =T+U.
  4. No cell should contain a formula referencing N, O, P, or Q.

Exit code: 0 = pass, 1 = violations found.
"""
import json, sys, re
from openpyxl import load_workbook

SESSION_FILE = ".audit-session.json"
with open(SESSION_FILE, encoding="utf-8") as f:
    session = json.load(f)

FT_PATH      = session["files"]["feuille_travail"]
data_start   = session.get("data_start", 4)
tot_paie_row = session.get("tot_paie_row", 178)

# Load WITHOUT data_only so we can read formulas
wb = load_workbook(FT_PATH)
ws = wb["Feuil2"]

violations = []

print("=" * 70)
print("eval_formulas — Formula Integrity Check")
print(f"  Checking rows {data_start}–{tot_paie_row-1}")
print("=" * 70)

BLANK_COLS  = {14, 15, 16, 17}   # N, O, P, Q
Y_COL_IDX   = 25
V_COL_IDX   = 22

# Regex for Y formula: should contain R, S, T, U, W, X (not N, O, P, Q)
Y_PATTERN  = re.compile(r"=R\d+\+S\d+\+T\d+\+U\d+\+W\d+\+X\d+", re.IGNORECASE)
V_PATTERN  = re.compile(r"=T\d+\+U\d+", re.IGNORECASE)
BAD_REFS   = re.compile(r"[=+\-\*\/]N\d+|[=+\-\*\/]O\d+|[=+\-\*\/]P\d+|[=+\-\*\/]Q\d+")

blank_violations = []
y_violations = []
v_violations = []
bad_ref_violations = []

for r in range(data_start, tot_paie_row):
    # Check N/O/P/Q are blank
    for col in BLANK_COLS:
        val = ws.cell(r, col).value
        if val not in (None, 0, ""):
            blank_violations.append(f"Row {r} Col {col}: expected blank, got '{val}'")

    # Check Y formula
    y_val = str(ws.cell(r, Y_COL_IDX).value or "")
    if not Y_PATTERN.search(y_val):
        y_violations.append(f"Row {r} Y: '{y_val}'")

    # Check V formula
    v_val = str(ws.cell(r, V_COL_IDX).value or "")
    if not V_PATTERN.search(v_val):
        v_violations.append(f"Row {r} V: '{v_val}'")

    # Check for bad N/O/P/Q references anywhere in the row
    for col in range(1, 30):
        val = str(ws.cell(r, col).value or "")
        if val.startswith("=") and BAD_REFS.search(val):
            bad_ref_violations.append(f"Row {r} Col {col}: '{val}' references N/O/P/Q")

# Report
if blank_violations:
    print(f"\n❌ N/O/P/Q not blank: {len(blank_violations)} violation(s)")
    for v in blank_violations[:10]:
        print(f"   • {v}")
else:
    print(f"✅ N/O/P/Q columns: all blank on {tot_paie_row - data_start} data rows")

if y_violations:
    print(f"\n❌ Y formula issues: {len(y_violations)} violation(s)")
    for v in y_violations[:10]:
        print(f"   • {v}")
else:
    print(f"✅ Y column: all formulas = R+S+T+U+W+X")

if v_violations:
    print(f"\n❌ V formula issues: {len(v_violations)} violation(s)")
    for v in v_violations[:10]:
        print(f"   • {v}")
else:
    print(f"✅ V column: all formulas = T+U")

if bad_ref_violations:
    print(f"\n❌ Stale N/O/P/Q references: {len(bad_ref_violations)} violation(s)")
    for v in bad_ref_violations[:10]:
        print(f"   • {v}")
else:
    print(f"✅ No stale N/O/P/Q formula references found")

wb.close()
all_violations = blank_violations + y_violations + v_violations + bad_ref_violations

print("\n" + "=" * 70)
if all_violations:
    print(f"❌ FAIL — {len(all_violations)} total formula violation(s)")
    sys.exit(1)
else:
    print("✅ PASS — All formula integrity checks passed")
    sys.exit(0)
