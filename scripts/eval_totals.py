"""
eval_totals.py — Independently recompute PAIE and COMPTA totals from source files
and compare to what is written in the FT-P-2 workbook TOTAL rows.

Exit code: 0 = all pass, 1 = mismatches found.
"""
import json, re, sys
import pandas as pd
from openpyxl import load_workbook

TOLERANCE = 1  # FCFA

SESSION_FILE = ".audit-session.json"
with open(SESSION_FILE, encoding="utf-8") as f:
    session = json.load(f)

FT_PATH = session["files"]["feuille_travail"]
BG_PATH = session["files"]["balance_generale"]
LP_PATH = ".livre_paie_pivot.csv"
CP_PATH = ".charges_patronales_pivot.csv"

# ── Recompute expected totals ───────────────────────────────────────────────────
# PAIE from pivot CSVs
lp = pd.read_csv(LP_PATH, dtype={"Matricule": str})
cp = pd.read_csv(CP_PATH, dtype={"Matricule": str})

expected_paie = {
    18: round(lp["SAL BRUT"].sum()),                                    # R
    19: round(cp.get("CNPS/P (Pension Vieillesse)", pd.Series([0])).sum()),  # S
    20: round(cp.get("CF/P (Crédit Foncier Patronal)", pd.Series([0])).sum()),  # T
    21: round(cp.get("FNE (Fond National Emploi)", pd.Series([0])).sum()),      # U
    23: round(cp.get("AF (Allocation Familiale)", pd.Series([0])).sum()),       # W
    24: round(cp.get("AT (Accident de Travail)", pd.Series([0])).sum()),        # X
}
expected_paie[22] = expected_paie[20] + expected_paie[21]  # V = T+U
expected_paie[25] = sum(expected_paie[c] for c in [18,19,20,21,23,24])  # Y

# COMPTA from Balance Générale
df_bg = pd.read_excel(BG_PATH, dtype={"Compte": str})
df_bg.columns = [str(c).strip() for c in df_bg.columns]
compte_col = df_bg.columns[0]
mvtd_col   = df_bg.columns[4] if len(df_bg.columns) > 4 else None
mvtc_col   = df_bg.columns[5] if len(df_bg.columns) > 5 else None
df_bg[compte_col] = df_bg[compte_col].astype(str).str.strip()
df_bg["NetSolde"] = pd.to_numeric(df_bg[mvtd_col], errors="coerce").fillna(0) - \
                    pd.to_numeric(df_bg[mvtc_col], errors="coerce").fillna(0)
account_map = {row[compte_col]: round(row["NetSolde"]) for _, row in df_bg.iterrows()}

GROUP_A = ["661110","661120","661130","661200","661210","661220","661300","661380","661410","661800","663101","663102","663410"]
GROUP_B_R_COL = {"664120": 19, "664110": 23, "664130": 24}
GROUP_C_R_COL = {"664380": 20}

expected_compta = {c: 0 for c in range(18, 26)}
for acct in GROUP_A:
    expected_compta[18] += account_map.get(acct, 0)
for acct, col in GROUP_B_R_COL.items():
    expected_compta[col] += account_map.get(acct, 0)
for acct, col in GROUP_C_R_COL.items():
    expected_compta[col] += account_map.get(acct, 0)
expected_compta[22] = expected_compta[20] + expected_compta[21]  # V
expected_compta[25] = sum(expected_compta[c] for c in [18,19,20,21,23,24])  # Y

# ── Read actual totals from workbook ───────────────────────────────────────────
wb = load_workbook(FT_PATH, read_only=True, data_only=True)
ws = wb["Feuil2"]

tot_paie_row   = session.get("tot_paie_row", 178)
tot_compta_row = session.get("tot_compta_row")

def read_row(ws, row_num):
    return {col: (ws.cell(row_num, col).value or 0) for col in range(18, 26)}

actual_paie   = read_row(ws, tot_paie_row)
actual_compta = read_row(ws, tot_compta_row) if tot_compta_row else {}
wb.close()

# ── Compare ─────────────────────────────────────────────────────────────────────
COL_NAMES = {18:"R(SAL BRUT)",19:"S(CNPS/P)",20:"T(CF/P)",21:"U(FNE)",22:"V(CF/P+FNE)",23:"W(AF)",24:"X(AT)",25:"Y(TOTAL)"}
failures = []

print("=" * 70)
print("eval_totals — Verification Report")
print("=" * 70)

print("\nPAIE (row %d):" % tot_paie_row)
for col in range(18, 26):
    exp = expected_paie.get(col, 0)
    act = actual_paie.get(col, 0)
    try:
        act_f = float(act)
    except (TypeError, ValueError):
        act_f = 0.0
    diff = abs(act_f - exp)
    status = "✅" if diff <= TOLERANCE else "❌"
    print(f"  {status} {COL_NAMES[col]:20s}: expected={exp:>20,.0f}  actual={act_f:>20,.0f}  diff={diff:,.0f}")
    if diff > TOLERANCE:
        failures.append(f"PAIE row {tot_paie_row} col {col}: expected {exp:,} got {act_f:,} diff {diff:,}")

if tot_compta_row:
    print(f"\nCOMPTA (row {tot_compta_row}):")
    for col in range(18, 26):
        exp = expected_compta.get(col, 0)
        act = actual_compta.get(col, 0)
        try:
            act_f = float(act)
        except (TypeError, ValueError):
            act_f = 0.0
        diff = abs(act_f - exp)
        status = "✅" if diff <= TOLERANCE else "❌"
        print(f"  {status} {COL_NAMES[col]:20s}: expected={exp:>20,.0f}  actual={act_f:>20,.0f}  diff={diff:,.0f}")
        if diff > TOLERANCE:
            failures.append(f"COMPTA row {tot_compta_row} col {col}: expected {exp:,} got {act_f:,} diff {diff:,}")

print("\n" + "=" * 70)
if failures:
    print(f"❌ FAIL — {len(failures)} mismatch(es) found")
    for f in failures:
        print(f"   • {f}")
    sys.exit(1)
else:
    print("✅ PASS — All column totals match source data")
    sys.exit(0)
