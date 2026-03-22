"""
parse_charges_patronales.py — Parse Charges Patronales CSV (non-standard Excel format).
Pivot per employee: CF/P (4100), FNE (4400), CNPS/P (4500), AF (4800), AT (4900).
Write to 'Extract Charges Patronal' sheet.
"""
import json, os, re, sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SESSION_FILE = ".audit-session.json"
with open(SESSION_FILE, encoding="utf-8") as f:
    session = json.load(f)

CSV_PATH = session["files"]["charges_patronales"]
FT_PATH  = session["files"]["feuille_travail"]

print(f"Parsing Charges Patronales: {CSV_PATH}")

# ── Custom CSV parser for ="value" format ──────────────────────────────────────
def parse_csv_value(raw):
    """Remove Excel text-forcing ="" wrapper and clean quotes."""
    s = raw.strip()
    s = re.sub(r'^=""?', '', s)
    s = re.sub(r'""?$', '', s)
    return s

def parse_amount(raw):
    """Parse French decimal comma amount: '87453","00' → 87453.00"""
    s = parse_csv_value(raw)
    s = s.replace('","', '.').replace(',', '.').replace(' ', '')
    try:
        return float(s)
    except ValueError:
        return 0.0

rows = []
with open(CSV_PATH, encoding="latin-1") as f:
    for line in f:
        line = line.rstrip('\n\r')
        if not line.strip():
            continue
        parts = line.split(';')
        if len(parts) < 6:
            continue
        try:
            matricule  = parse_csv_value(parts[0])
            nom        = parse_csv_value(parts[1])
            prenom     = parse_csv_value(parts[2]) if len(parts) > 2 else ""
            type_code  = parse_csv_value(parts[4]) if len(parts) > 4 else ""
            amount_raw = parts[5] if len(parts) > 5 else "0"
            amount     = parse_amount(amount_raw)
            if matricule and type_code:
                rows.append({
                    "Matricule": matricule,
                    "Nom": nom,
                    "Prenom": prenom,
                    "TypeCode": type_code,
                    "Montant": amount,
                })
        except Exception:
            continue

df = pd.DataFrame(rows)
print(f"  Raw rows parsed: {len(df)}")
print(f"  Type codes found: {df['TypeCode'].unique().tolist()[:10]}")

# Map codes to column names
CODE_MAP = {
    "4100": "CF/P (Crédit Foncier Patronal)",
    "4400": "FNE (Fond National Emploi)",
    "4500": "CNPS/P (Pension Vieillesse)",
    "4800": "AF (Allocation Familiale)",
    "4900": "AT (Accident de Travail)",
}
df = df[df["TypeCode"].isin(CODE_MAP.keys())]
df["ColName"] = df["TypeCode"].map(CODE_MAP)

# Pivot: one row per employee
pivot = df.pivot_table(
    index=["Matricule", "Nom", "Prenom"],
    columns="ColName",
    values="Montant",
    aggfunc="sum",
    fill_value=0,
).reset_index()
pivot.columns.name = None

# Ensure all charge columns exist
for col in CODE_MAP.values():
    if col not in pivot.columns:
        pivot[col] = 0

print(f"  Pivot: {len(pivot)} employees")

# ── Write Extract Charges Patronal sheet ───────────────────────────────────────
wb = load_workbook(FT_PATH)
if "Extract Charges Patronal" in wb.sheetnames:
    del wb["Extract Charges Patronal"]
ws = wb.create_sheet("Extract Charges Patronal")

BLUE_FILL = PatternFill("solid", start_color="1F4E79")
WHITE_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
DATA_FONT  = Font(name="Arial", size=10)
ALT_FILL   = PatternFill("solid", start_color="F5F8FC")
RGT = Alignment(horizontal="right", vertical="center")
LFT = Alignment(horizontal="left",  vertical="center")
NUM_FMT = "#,##0;(#,##0);\"-\""

def thin(): return Side(style="thin", color="D9D9D9")
BDR = Border(bottom=thin(), right=thin())

ordered_cols = [
    "Matricule", "Nom", "Prenom",
    "CF/P (Crédit Foncier Patronal)",
    "FNE (Fond National Emploi)",
    "CNPS/P (Pension Vieillesse)",
    "AF (Allocation Familiale)",
    "AT (Accident de Travail)",
]
existing = [c for c in ordered_cols if c in pivot.columns]
pivot = pivot[existing]

numeric_cols = set(ordered_cols[3:])

for col_idx, header in enumerate(existing, 1):
    c = ws.cell(row=1, column=col_idx, value=header)
    c.font = WHITE_FONT; c.fill = BLUE_FILL
    c.alignment = LFT; c.border = BDR

for r_idx, row in enumerate(pivot.itertuples(index=False), 2):
    alt = ALT_FILL if (r_idx % 2 == 0) else None
    for c_idx, (val, col_name) in enumerate(zip(row, existing), 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=val)
        cell.font = DATA_FONT; cell.border = BDR
        if alt: cell.fill = alt
        if col_name in numeric_cols:
            cell.number_format = NUM_FMT; cell.alignment = RGT
        else:
            cell.alignment = LFT

ws.auto_filter.ref = f"A1:{get_column_letter(len(existing))}1"
ws.freeze_panes = "A2"
ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 25
ws.column_dimensions["C"].width = 20
for i in range(4, len(existing)+1):
    ws.column_dimensions[get_column_letter(i)].width = 22

wb.save(FT_PATH)
print(f"✅ Extract Charges Patronal: {len(pivot)} employees written to '{FT_PATH}'")

# Persist pivot for reconciliation use
pivot.to_csv(".charges_patronales_pivot.csv", index=False, encoding="utf-8")

session.setdefault("extract_counts", {})["charges_patronales"] = len(pivot)
session.setdefault("steps_completed", [])
if "extract_charges_patronal" not in session["steps_completed"]:
    session["steps_completed"].append("extract_charges_patronal")
with open(SESSION_FILE, "w", encoding="utf-8") as f:
    json.dump(session, f, indent=2, ensure_ascii=False)
