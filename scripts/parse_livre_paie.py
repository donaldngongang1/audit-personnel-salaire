"""
parse_livre_paie.py — Parse Livre de Paie CSV (non-standard Excel format).
Pivot per employee: SAL BRUT (code BRUT).
Write to 'Extract LivrePaie' sheet.
"""
import json, re, sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SESSION_FILE = ".audit-session.json"
with open(SESSION_FILE, encoding="utf-8") as f:
    session = json.load(f)

CSV_PATH = session["files"]["livre_paie"]
FT_PATH  = session["files"]["feuille_travail"]

print(f"Parsing Livre de Paie: {CSV_PATH}")

def parse_csv_value(raw):
    s = raw.strip()
    s = re.sub(r'^=""?', '', s)
    s = re.sub(r'""?$', '', s)
    return s

def parse_amount(raw):
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

# Keep only BRUT rows
df_brut = df[df["TypeCode"].str.upper() == "BRUT"].copy()
print(f"  BRUT rows: {len(df_brut)}")

# Pivot: one row per employee
pivot = df_brut.groupby(["Matricule", "Nom", "Prenom"], as_index=False)["Montant"].sum()
pivot.rename(columns={"Montant": "SAL BRUT"}, inplace=True)
print(f"  Pivot: {len(pivot)} employees")

# ── Write Extract LivrePaie sheet ───────────────────────────────────────────────
wb = load_workbook(FT_PATH)
if "Extract LivrePaie" in wb.sheetnames:
    del wb["Extract LivrePaie"]
ws = wb.create_sheet("Extract LivrePaie")

BLUE_FILL = PatternFill("solid", start_color="1F4E79")
WHITE_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
DATA_FONT  = Font(name="Arial", size=10)
ALT_FILL   = PatternFill("solid", start_color="F5F8FC")
RGT = Alignment(horizontal="right", vertical="center")
LFT = Alignment(horizontal="left",  vertical="center")
NUM_FMT = "#,##0;(#,##0);\"-\""

def thin(): return Side(style="thin", color="D9D9D9")
BDR = Border(bottom=thin(), right=thin())

headers = ["Matricule", "Nom", "Prenom", "SAL BRUT"]

for col_idx, header in enumerate(headers, 1):
    c = ws.cell(row=1, column=col_idx, value=header)
    c.font = WHITE_FONT; c.fill = BLUE_FILL
    c.alignment = LFT; c.border = BDR

for r_idx, row in enumerate(pivot.itertuples(index=False), 2):
    alt = ALT_FILL if (r_idx % 2 == 0) else None
    vals = [row.Matricule, row.Nom, row.Prenom, row._3]  # SAL BRUT is 4th field
    for c_idx, val in enumerate(vals, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=val)
        cell.font = DATA_FONT; cell.border = BDR
        if alt: cell.fill = alt
        if c_idx == 4:
            cell.number_format = NUM_FMT; cell.alignment = RGT
        else:
            cell.alignment = LFT

ws.auto_filter.ref = "A1:D1"
ws.freeze_panes = "A2"
ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 25
ws.column_dimensions["C"].width = 20
ws.column_dimensions["D"].width = 20

wb.save(FT_PATH)
print(f"✅ Extract LivrePaie: {len(pivot)} employees written to '{FT_PATH}'")

# Persist pivot for reconciliation
pivot.to_csv(".livre_paie_pivot.csv", index=False, encoding="utf-8")

session.setdefault("extract_counts", {})["livre_paie"] = len(pivot)
session.setdefault("steps_completed", [])
if "extract_livre_paie" not in session["steps_completed"]:
    session["steps_completed"].append("extract_livre_paie")
with open(SESSION_FILE, "w", encoding="utf-8") as f:
    json.dump(session, f, indent=2, ensure_ascii=False)
