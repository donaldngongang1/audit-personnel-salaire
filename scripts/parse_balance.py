"""
parse_balance.py — Parse Balance Générale (.xlsx) and write Extract Balance sheet.
Filters for accounts 661–663 (Charges du personnel).
Net Solde = MvtDebit − MvtCredit.
"""
import json, os, sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Load session ────────────────────────────────────────────────────────────────
SESSION_FILE = ".audit-session.json"
with open(SESSION_FILE, encoding="utf-8") as f:
    session = json.load(f)

BG_PATH = session["files"]["balance_generale"]
FT_PATH = session["files"]["feuille_travail"]

print(f"Parsing Balance Générale: {BG_PATH}")

# ── Read Balance Générale ───────────────────────────────────────────────────────
df = pd.read_excel(BG_PATH, dtype={"Compte": str})

# Detect column names (may vary; normalize)
df.columns = [str(c).strip() for c in df.columns]

# Find account and movement columns
compte_col = next((c for c in df.columns if c.lower() in ["compte", "account", "n° compte"]), df.columns[0])
libelle_col = next((c for c in df.columns if "lib" in c.lower()), df.columns[1])
mvt_debit_col = next((c for c in df.columns if "mvtdeb" in c.lower().replace(" ", "") or "débit" in c.lower() and "mvt" in c.lower()), None)
mvt_credit_col = next((c for c in df.columns if "mvtcre" in c.lower().replace(" ", "") or "crédit" in c.lower() and "mvt" in c.lower()), None)

# Fallback to positional (standard 8-column format):
# Compte | Libellé | SolDebitOuv | SolCreditOuv | MvtDebit | MvtCredit | SolDebitClo | SolCreditClo
if mvt_debit_col is None and len(df.columns) >= 5:
    mvt_debit_col = df.columns[4]
if mvt_credit_col is None and len(df.columns) >= 6:
    mvt_credit_col = df.columns[5]

print(f"  Columns: {list(df.columns)}")
print(f"  Account col: {compte_col}, MvtDebit: {mvt_debit_col}, MvtCredit: {mvt_credit_col}")

# ── Filter accounts 661–663 ─────────────────────────────────────────────────────
mask = df[compte_col].astype(str).str.match(r'^66[123]')
df_filtered = df[mask].copy()

# Compute net solde
df_filtered["MvtDebit_"] = pd.to_numeric(df_filtered[mvt_debit_col], errors="coerce").fillna(0)
df_filtered["MvtCredit_"] = pd.to_numeric(df_filtered[mvt_credit_col], errors="coerce").fillna(0)
df_filtered["NetSolde"] = df_filtered["MvtDebit_"] - df_filtered["MvtCredit_"]

print(f"  Filtered: {len(df_filtered)} rows (accounts 661–663)")

# ── Write Extract Balance sheet ─────────────────────────────────────────────────
wb = load_workbook(FT_PATH)
if "Extract Balance" in wb.sheetnames:
    del wb["Extract Balance"]
ws = wb.create_sheet("Extract Balance")

BLUE_FILL = PatternFill("solid", start_color="1F4E79")
WHITE_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
DATA_FONT  = Font(name="Arial", size=10)
ALT_FILL   = PatternFill("solid", start_color="F5F8FC")
RGT = Alignment(horizontal="right", vertical="center")
LFT = Alignment(horizontal="left",  vertical="center")
NUM_FMT = "#,##0;(#,##0);\"-\""

def thin(): return Side(style="thin", color="D9D9D9")
BDR = Border(bottom=thin(), right=thin())

headers = ["Compte", "Libellé", "MvtDebit", "MvtCredit", "NetSolde"]

for col_idx, header in enumerate(headers, 1):
    c = ws.cell(row=1, column=col_idx, value=header)
    c.font = WHITE_FONT; c.fill = BLUE_FILL; c.alignment = LFT
    c.border = BDR

output_cols = [compte_col, libelle_col, mvt_debit_col, mvt_credit_col]
df_out = df_filtered[output_cols].copy()
df_out.columns = ["Compte", "Libellé", "MvtDebit", "MvtCredit"]
df_out["NetSolde"] = df_filtered["NetSolde"].values

for r_idx, row in enumerate(df_out.itertuples(index=False), 2):
    alt = ALT_FILL if (r_idx % 2 == 0) else None
    for c_idx, val in enumerate(row, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=val)
        cell.font = DATA_FONT; cell.border = BDR
        if alt: cell.fill = alt
        if c_idx in [3, 4, 5]:
            cell.number_format = NUM_FMT
            cell.alignment = RGT
        else:
            cell.alignment = LFT

# Auto-filter and freeze
ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
ws.freeze_panes = "A2"

# Column widths
ws.column_dimensions["A"].width = 12
ws.column_dimensions["B"].width = 40
for col in ["C", "D", "E"]:
    ws.column_dimensions[col].width = 18

wb.save(FT_PATH)
print(f"✅ Extract Balance: {len(df_out)} rows written to '{FT_PATH}'")

# Update session
session.setdefault("extract_counts", {})["balance"] = len(df_out)
session.setdefault("steps_completed", [])
if "extract_balance" not in session["steps_completed"]:
    session["steps_completed"].append("extract_balance")
with open(SESSION_FILE, "w", encoding="utf-8") as f:
    json.dump(session, f, indent=2, ensure_ascii=False)
